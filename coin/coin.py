import hashlib
import json
import sys
import time
import uuid
from decimal import Decimal, InvalidOperation, ROUND_DOWN
from urllib.parse import urlencode

import jwt
import pandas as pd
import requests
from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


def check_scalping_signal(df):
    if len(df) < 21:
        return False
    df = df.copy()
    df["EMA20"] = df["trade_price"].ewm(span=20).mean()
    df["Volume_Increase"] = df["candle_acc_trade_volume"] > df["candle_acc_trade_volume"].shift(1)
    last = df.iloc[-1]
    prev = df.iloc[-2]
    return (
        last["trade_price"] > last["EMA20"]
        and prev["trade_price"] <= prev["EMA20"]
        and bool(last["Volume_Increase"])
    )


def backtest_scalping_strategy(market="KRW-BTC", count=200, initial_fund=1_000_000):
    url = "https://api.upbit.com/v1/candles/minutes/1"
    response = requests.get(url, params={"market": market, "count": count}, timeout=10)
    response.raise_for_status()
    raw = response.json()
    if not raw:
        return [], float(initial_fund)

    df = pd.DataFrame(raw)[::-1]
    df["candle_date_time_kst"] = pd.to_datetime(df["candle_date_time_kst"])
    df.set_index("candle_date_time_kst", inplace=True)
    df = df[["trade_price", "candle_acc_trade_volume"]]

    signals = []
    holding = False
    entry_price = 0.0
    balance = float(initial_fund)

    for i in range(20, len(df)):
        sliced = df.iloc[i - 20 : i + 1]
        price = float(sliced.iloc[-1]["trade_price"])
        if not holding and check_scalping_signal(sliced):
            entry_price = price
            signals.append((sliced.index[-1], "BUY", price))
            holding = True
        elif holding:
            profit = (price - entry_price) / entry_price
            if profit >= 0.005 or profit <= -0.003:
                signals.append((sliced.index[-1], "SELL", price))
                balance *= 1 + profit
                holding = False

    return signals, balance


class ApiLoginDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Upbit Login")
        self.setFixedSize(430, 170)

        layout = QVBoxLayout()
        form = QFormLayout()

        self.access_input = QLineEdit()
        self.access_input.setPlaceholderText("Access Key")
        form.addRow("Access Key", self.access_input)

        self.secret_input = QLineEdit()
        self.secret_input.setEchoMode(QLineEdit.Password)
        self.secret_input.setPlaceholderText("Secret Key")
        form.addRow("Secret Key", self.secret_input)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Login")
        buttons.accepted.connect(self._submit)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def _submit(self):
        if not self.access_input.text().strip() or not self.secret_input.text().strip():
            QMessageBox.warning(self, "Input Error", "Access Key and Secret Key are required.")
            return
        self.accept()

    def get_credentials(self):
        return self.access_input.text().strip(), self.secret_input.text().strip()


class TradingBot(QWidget):
    def __init__(self, access_key, secret_key):
        super().__init__()
        self.setWindowTitle("Coin Trading Bot")
        self.setFixedSize(900, 700)

        self.access_key = access_key
        self.secret_key = secret_key
        self.server_url = "https://api.upbit.com"

        self.market_names = {}
        self.holding = False
        self.entry_price = None
        self.last_status_log = 0.0

        self.load_coin_names()
        self._build_ui()

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_price_info)

    def _build_ui(self):
        self.tabs = QTabWidget()
        self.trade_tab = QWidget()
        self.account_tab = QWidget()
        self.backtest_tab = QWidget()
        self.tabs.addTab(self.trade_tab, "Trade")
        self.tabs.addTab(self.account_tab, "Account")
        self.tabs.addTab(self.backtest_tab, "Backtest")

        self._init_trade_tab()
        self._init_account_tab()
        self._init_backtest_tab()

        layout = QVBoxLayout()
        layout.addWidget(self.tabs)
        self.setLayout(layout)

    def _init_trade_tab(self):
        layout = QVBoxLayout()

        self.market_selector = QComboBox()
        markets = sorted(self.market_names.keys())
        for market in markets:
            display = f"{self.market_names.get(market, market)} ({market})"
            self.market_selector.addItem(display, market)
        default_index = self.market_selector.findData("KRW-BTC")
        if default_index >= 0:
            self.market_selector.setCurrentIndex(default_index)

        self.trigger_input = QLineEdit()
        self.trigger_input.setPlaceholderText("Optional trigger price")
        self.amount_input = QLineEdit()
        self.amount_input.setPlaceholderText("KRW amount, e.g. 10000")

        self.radio_market = QRadioButton("Market Buy")
        self.radio_limit = QRadioButton("Limit Buy")
        self.radio_market.setChecked(True)

        self.start_button = QPushButton("Start")
        self.start_button.clicked.connect(self.start_bot)
        self.stop_button = QPushButton("Stop")
        self.stop_button.clicked.connect(self.stop_bot)
        self.stop_button.setEnabled(False)
        self.manual_order_button = QPushButton("Manual Buy")
        self.manual_order_button.clicked.connect(self.manual_place_order)
        self.manual_sell_button = QPushButton("Manual Sell")
        self.manual_sell_button.clicked.connect(self.manual_sell_order)

        self.price_label = QLabel("Price: -")
        self.volume_label = QLabel("24h Volume: -")
        self.change_label = QLabel("Daily Change: -")
        self.status_label = QLabel("Status: Idle")
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)

        buttons = QHBoxLayout()
        buttons.addWidget(self.start_button)
        buttons.addWidget(self.stop_button)
        buttons.addWidget(self.manual_order_button)
        buttons.addWidget(self.manual_sell_button)

        radios = QHBoxLayout()
        radios.addWidget(self.radio_market)
        radios.addWidget(self.radio_limit)

        layout.addWidget(QLabel("Market"))
        layout.addWidget(self.market_selector)
        layout.addWidget(QLabel("Trigger Price (optional)"))
        layout.addWidget(self.trigger_input)
        layout.addWidget(QLabel("Buy Amount (KRW)"))
        layout.addWidget(self.amount_input)
        layout.addLayout(radios)
        layout.addLayout(buttons)
        layout.addWidget(self.price_label)
        layout.addWidget(self.volume_label)
        layout.addWidget(self.change_label)
        layout.addWidget(self.status_label)
        layout.addWidget(QLabel("Log"))
        layout.addWidget(self.log_output)
        self.trade_tab.setLayout(layout)

    def _init_account_tab(self):
        layout = QVBoxLayout()
        self.balance_label = QLabel("KRW Balance: -")
        self.holdings_label = QLabel("Holdings: -")
        self.eval_label = QLabel("Total Evaluation: -")
        self.profit_label = QLabel("Profit Rate: -")
        refresh_button = QPushButton("Refresh")
        refresh_button.clicked.connect(self.get_balance)
        layout.addWidget(self.balance_label)
        layout.addWidget(self.holdings_label)
        layout.addWidget(self.eval_label)
        layout.addWidget(self.profit_label)
        layout.addWidget(refresh_button)
        self.account_tab.setLayout(layout)

    def _init_backtest_tab(self):
        layout = QVBoxLayout()
        self.initial_fund_input = QLineEdit()
        self.initial_fund_input.setPlaceholderText("1000000")
        self.backtest_button = QPushButton("Run Backtest")
        self.backtest_button.clicked.connect(self.run_backtest)
        self.backtest_output = QTextEdit()
        self.backtest_output.setReadOnly(True)
        layout.addWidget(QLabel("Initial Fund (KRW)"))
        layout.addWidget(self.initial_fund_input)
        layout.addWidget(self.backtest_button)
        layout.addWidget(self.backtest_output)
        self.backtest_tab.setLayout(layout)

    def log(self, message):
        self.log_output.append(message)

    def load_coin_names(self):
        try:
            res = requests.get(
                f"{self.server_url}/v1/market/all",
                params={"isDetails": "false"},
                timeout=10,
            )
            res.raise_for_status()
            data = res.json()
            self.market_names = {x["market"]: x["korean_name"] for x in data if x["market"].startswith("KRW-")}
        except Exception:
            self.market_names = {"KRW-BTC": "Bitcoin"}

    def selected_market(self):
        market = self.market_selector.currentData()
        if not market:
            raise ValueError("No market selected.")
        return market

    def start_bot(self):
        self.update_price_info()
        self.timer.start(5000)
        self.status_label.setText("Status: Running")
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.get_balance()
        self.log("Bot started.")

    def stop_bot(self):
        self.timer.stop()
        self.status_label.setText("Status: Stopped")
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.log("Bot stopped.")

    def run_backtest(self):
        try:
            market = self.selected_market()
            initial = int(self.initial_fund_input.text()) if self.initial_fund_input.text().isdigit() else 1_000_000
            self.backtest_output.append("Backtest started...")
            signals, final_balance = backtest_scalping_strategy(market=market, initial_fund=initial)
            if not signals:
                self.backtest_output.append("No signals found.")
            else:
                for at_time, side, price in signals:
                    self.backtest_output.append(f"{at_time.strftime('%Y-%m-%d %H:%M')} {side} @ {price:,.0f}")
            rate = (final_balance - initial) / initial * 100 if initial > 0 else 0
            self.backtest_output.append(f"Final balance: {final_balance:,.0f} KRW ({rate:+.2f}%)\n")
        except Exception as exc:
            self.backtest_output.append(f"Backtest error: {exc}")

    def update_price_info(self):
        try:
            market = self.selected_market()
            ticker = requests.get(
                f"{self.server_url}/v1/ticker",
                params={"markets": market},
                timeout=10,
            )
            ticker.raise_for_status()
            res = ticker.json()[0]
            current_price = float(res["trade_price"])
            prev_price = float(res["prev_closing_price"])
            rate = ((current_price - prev_price) / prev_price * 100) if prev_price else 0

            self.price_label.setText(f"Price: {current_price:,.8f} KRW")
            self.volume_label.setText(f"24h Volume: {float(res['acc_trade_price_24h']):,.0f} KRW")
            self.change_label.setText(f"Daily Change: {rate:+.2f}%")

            trigger_text = self.trigger_input.text().strip()
            if trigger_text and not self.holding:
                try:
                    trigger = float(trigger_text)
                    if current_price <= trigger:
                        self.log(f"Trigger hit ({trigger:,.0f}). Trying buy.")
                        self.place_buy_order(current_price)
                except ValueError:
                    self.log("Trigger price ignored (invalid number).")

            if self.holding and self.entry_price:
                pnl = (current_price - self.entry_price) / self.entry_price
                now = time.time()
                if now - self.last_status_log >= 10:
                    self.log(f"Current PnL: {pnl:+.2%}")
                    self.last_status_log = now
                if pnl >= 0.01 or pnl <= -0.005:
                    self.log(f"Auto sell condition met ({pnl:+.2%}).")
                    self.place_sell_order(current_price)
        except Exception as exc:
            self.price_label.setText("Price: failed")
            self.log(f"Price update failed: {exc}")

    def manual_place_order(self):
        try:
            market = self.selected_market()
            ticker = requests.get(
                f"{self.server_url}/v1/ticker",
                params={"markets": market},
                timeout=10,
            )
            ticker.raise_for_status()
            current_price = float(ticker.json()[0]["trade_price"])
            self.place_buy_order(current_price)
        except Exception as exc:
            self.log(f"Manual buy error: {exc}")

    def manual_sell_order(self):
        try:
            market = self.selected_market()
            ticker = requests.get(
                f"{self.server_url}/v1/ticker",
                params={"markets": market},
                timeout=10,
            )
            ticker.raise_for_status()
            current_price = float(ticker.json()[0]["trade_price"])
            self.place_sell_order(current_price)
        except Exception as exc:
            self.log(f"Manual sell error: {exc}")

    def adjust_price(self, price):
        steps = [
            (0, 0.001, 0.0001),
            (0.001, 0.01, 0.0001),
            (0.01, 0.1, 0.001),
            (0.1, 1, 0.01),
            (1, 10, 0.1),
            (10, 100, 1),
            (100, 1000, 5),
            (1000, 10000, 10),
            (10000, 100000, 50),
            (100000, 500000, 100),
            (500000, 1000000, 500),
            (1000000, float("inf"), 1000),
        ]
        for low, high, unit in steps:
            if low <= price < high:
                return float((Decimal(str(price)) // Decimal(str(unit))) * Decimal(str(unit)))
        return price

    def place_buy_order(self, current_price):
        if self.holding:
            self.log("Buy skipped: already holding.")
            return

        market = self.selected_market()
        amount_text = self.amount_input.text().strip()
        try:
            amount = Decimal(amount_text)
            if amount <= 0:
                raise InvalidOperation
        except (InvalidOperation, ValueError):
            self.log("Invalid buy amount.")
            return

        if amount < Decimal("5000"):
            self.log("Buy amount must be at least 5,000 KRW.")
            return

        order_type = "price" if self.radio_market.isChecked() else "limit"
        body = {"market": market, "side": "bid"}

        if order_type == "price":
            body["ord_type"] = "price"
            body["price"] = str(amount.quantize(Decimal("1"), rounding=ROUND_DOWN))
        else:
            adjusted_price = self.adjust_price(current_price)
            volume = (amount / Decimal(str(adjusted_price))).quantize(
                Decimal("0.00000001"), rounding=ROUND_DOWN
            )
            body["ord_type"] = "limit"
            body["price"] = str(adjusted_price)
            body["volume"] = str(volume)

        try:
            order = self.auth_request("POST", "/v1/orders", body)
            self.entry_price = current_price
            self.holding = True
            self.status_label.setText("Status: Bought")
            self.log(f"Buy success: {order.get('uuid', '-')}")
        except Exception as exc:
            self.status_label.setText("Status: Buy failed")
            self.log(f"Buy failed: {exc}")

    def place_sell_order(self, current_price):
        if not self.holding:
            self.log("Sell skipped: no holding state.")
            return

        market = self.selected_market()
        coin = market.split("-")[1]

        try:
            accounts = self.get_accounts()
            coin_account = next((a for a in accounts if a["currency"] == coin), None)
            volume = Decimal(coin_account["balance"]) if coin_account else Decimal("0")
            if volume <= 0:
                self.log("Sell failed: no coin balance.")
                return

            body = {
                "market": market,
                "side": "ask",
                "ord_type": "market",
                "volume": str(volume.quantize(Decimal("0.00000001"), rounding=ROUND_DOWN)),
            }
            order = self.auth_request("POST", "/v1/orders", body)
            pnl = (
                (current_price - self.entry_price) / self.entry_price * 100
                if self.entry_price
                else 0
            )
            self.holding = False
            self.entry_price = None
            self.status_label.setText("Status: Sold")
            self.log(f"Sell success: {order.get('uuid', '-')}, PnL {pnl:+.2f}%")
        except Exception as exc:
            self.status_label.setText("Status: Sell failed")
            self.log(f"Sell failed: {exc}")

    def auth_headers(self, query):
        query_string = urlencode(query).encode()
        query_hash = hashlib.sha512(query_string).hexdigest()
        payload = {
            "access_key": self.access_key,
            "nonce": str(uuid.uuid4()),
            "query_hash": query_hash,
            "query_hash_alg": "SHA512",
        }
        token = jwt.encode(payload, self.secret_key, algorithm="HS256")
        return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    def auth_request(self, method, path, body):
        headers = self.auth_headers(body)
        response = requests.request(
            method,
            f"{self.server_url}{path}",
            json=body,
            headers=headers,
            timeout=10,
        )
        data = response.json()
        if response.status_code >= 400:
            message = data.get("error", {}).get("message", f"HTTP {response.status_code}")
            raise RuntimeError(message)
        return data

    def get_accounts(self):
        payload = {"access_key": self.access_key, "nonce": str(uuid.uuid4())}
        token = jwt.encode(payload, self.secret_key, algorithm="HS256")
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(f"{self.server_url}/v1/accounts", headers=headers, timeout=10)
        data = response.json()
        if response.status_code >= 400:
            message = data.get("error", {}).get("message", f"HTTP {response.status_code}")
            raise RuntimeError(message)
        return data

    def get_balance(self):
        try:
            accounts = self.get_accounts()
            krw = next((x for x in accounts if x["currency"] == "KRW"), None)
            krw_balance = float(krw["balance"]) if krw else 0.0
            self.balance_label.setText(f"KRW Balance: {krw_balance:,.0f} KRW")

            coins = [x for x in accounts if x["currency"] != "KRW" and float(x["balance"]) > 0]
            summaries = []
            total_eval = 0.0
            total_cost = 0.0

            for coin in coins:
                market = f"KRW-{coin['currency']}"
                amount = float(coin["balance"])
                avg_price = float(coin.get("avg_buy_price", 0))
                try:
                    ticker = requests.get(
                        f"{self.server_url}/v1/ticker",
                        params={"markets": market},
                        timeout=5,
                    )
                    ticker.raise_for_status()
                    now_price = float(ticker.json()[0]["trade_price"])
                except Exception:
                    now_price = 0

                eval_value = now_price * amount
                cost_value = avg_price * amount
                total_eval += eval_value
                total_cost += cost_value
                name = self.market_names.get(market, coin["currency"])
                summaries.append(f"{name}({coin['currency']}): {amount:.8f}")

            self.holdings_label.setText("Holdings: " + (", ".join(summaries) if summaries else "-"))
            self.eval_label.setText(f"Total Evaluation: {total_eval:,.0f} KRW")
            profit_rate = ((total_eval - total_cost) / total_cost * 100) if total_cost else 0
            self.profit_label.setText(f"Profit Rate: {profit_rate:+.2f}%")
        except Exception as exc:
            self.balance_label.setText("KRW Balance: failed")
            self.log(f"Balance query failed: {exc}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    login = ApiLoginDialog()
    if login.exec_() != QDialog.Accepted:
        sys.exit(0)

    access_key, secret_key = login.get_credentials()
    window = TradingBot(access_key, secret_key)
    window.show()
    sys.exit(app.exec_())
