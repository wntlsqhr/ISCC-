import hashlib
import json
import os
import sys
import time
import uuid
from decimal import Decimal, InvalidOperation, ROUND_DOWN, getcontext
from urllib.parse import urlencode

import jwt
import requests
from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

getcontext().prec = 18


class ApiKeyLoginDialog(QDialog):
    def __init__(self, prefill_access="", prefill_secret=""):
        super().__init__()
        self.setWindowTitle("Upbit Login")
        self.setFixedSize(420, 160)

        layout = QVBoxLayout()
        form = QFormLayout()

        self.access_input = QLineEdit(prefill_access)
        self.access_input.setPlaceholderText("Access Key")
        form.addRow("Access Key", self.access_input)

        self.secret_input = QLineEdit(prefill_secret)
        self.secret_input.setEchoMode(QLineEdit.Password)
        self.secret_input.setPlaceholderText("Secret Key")
        form.addRow("Secret Key", self.secret_input)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Login")
        buttons.accepted.connect(self._validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

    def _validate_and_accept(self):
        if not self.access_input.text().strip() or not self.secret_input.text().strip():
            QMessageBox.warning(self, "Input Error", "Access Key and Secret Key are required.")
            return
        self.accept()

    def get_keys(self):
        return self.access_input.text().strip(), self.secret_input.text().strip()


class UpbitTrader(QWidget):
    def __init__(self, access_key, secret_key):
        super().__init__()
        self.setWindowTitle("Upbit Auto Trader")
        self.setFixedSize(720, 680)

        self.server_url = "https://api.upbit.com"
        self.access_key = access_key.strip()
        self.secret_key = secret_key.strip()

        self.market = None
        self.entry_price = None
        self.monitoring = False
        self.last_log_time = 0.0
        self._credential_warned = False

        self._build_ui()

        self.auto_sell_timer = QTimer(self)
        self.auto_sell_timer.timeout.connect(self.check_auto_sell)
        self.auto_sell_timer.start(1000)

        self.balance_timer = QTimer(self)
        self.balance_timer.timeout.connect(self.update_balances)
        self.balance_timer.start(2000)

    def _build_ui(self):
        root = QVBoxLayout()

        buy_group = QGroupBox("Buy")
        buy_form = QFormLayout()

        self.market_input = QLineEdit()
        self.market_input.setPlaceholderText("KRW-BTC")
        buy_form.addRow("Market", self.market_input)

        self.amount_input = QLineEdit()
        self.amount_input.setPlaceholderText("5000")
        buy_form.addRow("Buy KRW", self.amount_input)

        buy_group.setLayout(buy_form)
        root.addWidget(buy_group)

        balance_group = QGroupBox("Balances")
        balance_layout = QVBoxLayout()
        self.krw_balance_label = QLabel("KRW: -")
        self.coin_balance_label = QLabel("Coin: -")
        balance_layout.addWidget(self.krw_balance_label)
        balance_layout.addWidget(self.coin_balance_label)
        balance_group.setLayout(balance_layout)
        root.addWidget(balance_group)

        trade_group = QGroupBox("Trade / Auto Sell")
        trade_form = QFormLayout()
        self.take_profit_input = QLineEdit()
        self.take_profit_input.setPlaceholderText("5")
        trade_form.addRow("Take Profit (%)", self.take_profit_input)

        self.stop_loss_input = QLineEdit()
        self.stop_loss_input.setPlaceholderText("-3")
        trade_form.addRow("Stop Loss (%)", self.stop_loss_input)

        button_row = QHBoxLayout()
        self.buy_button = QPushButton("Buy Market")
        self.buy_button.clicked.connect(self.buy)
        self.sell_button = QPushButton("Sell Market")
        self.sell_button.clicked.connect(self.sell)
        self.auto_sell_button = QPushButton("Start Auto Sell")
        self.auto_sell_button.clicked.connect(self.toggle_auto_sell)
        button_row.addWidget(self.buy_button)
        button_row.addWidget(self.sell_button)
        button_row.addWidget(self.auto_sell_button)
        trade_form.addRow(button_row)

        trade_group.setLayout(trade_form)
        root.addWidget(trade_group)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        root.addWidget(self.log_output)

        self.setLayout(root)

    def log(self, message):
        self.log_output.append(message)

    def ensure_credentials(self, interactive=False):
        if self.access_key and self.secret_key:
            self._credential_warned = False
            return True
        if not self._credential_warned:
            self.log("Missing API keys: set UPBIT_ACCESS_KEY and UPBIT_SECRET_KEY.")
            self._credential_warned = True
        if interactive:
            QMessageBox.warning(
                self,
                "Credentials Required",
                "Set UPBIT_ACCESS_KEY and UPBIT_SECRET_KEY environment variables.",
            )
        return False

    def ensure_market(self, market):
        if not market:
            return False
        parts = market.split("-")
        return len(parts) == 2 and all(parts)

    def update_balances(self):
        if not self.ensure_credentials(interactive=False):
            return
        try:
            accounts = self.get_accounts()
            if accounts is None:
                return

            coin_code = self.market.split("-")[1] if self.market and "-" in self.market else None
            coin_balance = Decimal("0")

            for balance in accounts:
                if balance.get("currency") == "KRW":
                    krw = Decimal(balance.get("balance", "0"))
                    self.krw_balance_label.setText(f"KRW: {krw:,.0f}")
                if coin_code and balance.get("currency") == coin_code:
                    coin_balance = Decimal(balance.get("balance", "0"))

            if coin_code:
                self.coin_balance_label.setText(f"{coin_code}: {coin_balance:,.8f}")
            else:
                self.coin_balance_label.setText("Coin: -")
        except Exception as exc:
            self.log(f"Balance update error: {exc}")

    def buy(self):
        if not self.ensure_credentials(interactive=True):
            return

        market = self.market_input.text().strip().upper()
        amount_text = self.amount_input.text().strip()

        if not self.ensure_market(market):
            QMessageBox.warning(self, "Input Error", "Invalid market format. Example: KRW-BTC")
            return

        try:
            amount = Decimal(amount_text)
            if amount <= 0:
                raise InvalidOperation
        except (InvalidOperation, ValueError):
            QMessageBox.warning(self, "Input Error", "Buy amount must be a positive number.")
            return

        order_price = amount.quantize(Decimal("1"), rounding=ROUND_DOWN)
        if order_price < Decimal("5000"):
            QMessageBox.warning(self, "Input Error", "Upbit minimum KRW order is 5,000.")
            return

        self.market = market
        query = {
            "market": market,
            "side": "bid",
            "ord_type": "price",
            "price": str(order_price),
        }

        try:
            result = self.auth_request("POST", "/v1/orders", query)
            if not result:
                return

            self.log(f"Buy submitted: {json.dumps(result, ensure_ascii=False)}")
            order_uuid = result.get("uuid")
            fill_price = self.get_order_fill_price(order_uuid)
            if fill_price is not None:
                self.entry_price = fill_price
                self.log(f"Entry price: {self.entry_price:,.2f}")
            else:
                self.entry_price = None
                self.log("Buy placed but fill price was not confirmed.")
        except Exception as exc:
            self.log(f"Buy error: {exc}")

    def sell(self):
        if not self.ensure_credentials(interactive=True):
            return
        if not self.ensure_market(self.market):
            self.log("No market selected. Buy first or set a valid market.")
            return

        coin_code = self.market.split("-")[1]
        volume = self.get_balance(coin_code)
        if volume <= Decimal("0"):
            self.log("Sell skipped: no coin balance.")
            return

        query = {
            "market": self.market,
            "side": "ask",
            "ord_type": "market",
            "volume": str(volume.quantize(Decimal("0.00000001"), rounding=ROUND_DOWN)),
        }

        try:
            result = self.auth_request("POST", "/v1/orders", query)
            if not result:
                return

            self.log(f"Sell submitted: {json.dumps(result, ensure_ascii=False)}")
            order_uuid = result.get("uuid")
            sell_price = self.get_order_fill_price(order_uuid)

            if self.entry_price is not None and sell_price is not None:
                profit_rate = (sell_price - self.entry_price) / self.entry_price * Decimal("100")
                self.log(f"Exit price: {sell_price:,.2f}, PnL: {profit_rate:+.2f}%")
            else:
                self.log("Sell placed but fill price/PnL could not be calculated.")

            self.monitoring = False
            self.auto_sell_button.setText("Start Auto Sell")
            self.entry_price = None
        except Exception as exc:
            self.log(f"Sell error: {exc}")

    def toggle_auto_sell(self):
        if self.entry_price is None:
            self.log("Auto sell requires an entry price. Buy first.")
            return
        self.monitoring = not self.monitoring
        self.auto_sell_button.setText("Stop Auto Sell" if self.monitoring else "Start Auto Sell")
        self.log("Auto sell enabled." if self.monitoring else "Auto sell disabled.")

    def check_auto_sell(self):
        if not self.monitoring or self.entry_price is None or not self.market:
            return

        current_price = self.get_current_price(self.market)
        if current_price is None:
            return

        try:
            take_profit = Decimal(self.take_profit_input.text().strip())
            stop_loss = Decimal(self.stop_loss_input.text().strip())
        except (InvalidOperation, ValueError):
            self.log("Auto sell thresholds must be numeric.")
            return

        profit_rate = (current_price - self.entry_price) / self.entry_price * Decimal("100")

        now = time.time()
        if now - self.last_log_time >= 10:
            self.log(f"Current PnL: {profit_rate:+.2f}%")
            self.last_log_time = now

        if profit_rate >= take_profit:
            self.log(f"Take-profit triggered at {profit_rate:+.2f}%")
            self.sell()
        elif profit_rate <= stop_loss:
            self.log(f"Stop-loss triggered at {profit_rate:+.2f}%")
            self.sell()

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
        return {"Authorization": f"Bearer {token}"}

    def auth_request(self, method, path, query):
        headers = self.auth_headers(query)
        url = f"{self.server_url}{path}"
        response = requests.request(method, url, params=query, headers=headers, timeout=10)
        data = response.json()
        if response.status_code >= 400:
            message = data.get("error", {}).get("message", f"HTTP {response.status_code}")
            self.log(f"Upbit API error: {message}")
            return None
        return data

    def get_accounts(self):
        payload = {"access_key": self.access_key, "nonce": str(uuid.uuid4())}
        token = jwt.encode(payload, self.secret_key, algorithm="HS256")
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(f"{self.server_url}/v1/accounts", headers=headers, timeout=10)
        data = response.json()
        if response.status_code >= 400:
            message = data.get("error", {}).get("message", f"HTTP {response.status_code}")
            self.log(f"Accounts API error: {message}")
            return None
        return data

    def get_current_price(self, market):
        try:
            response = requests.get(
                f"{self.server_url}/v1/ticker",
                params={"markets": market},
                timeout=5,
            )
            data = response.json()
            if not data:
                return None
            return Decimal(str(data[0]["trade_price"]))
        except Exception:
            return None

    def get_balance(self, currency):
        accounts = self.get_accounts()
        if accounts is None:
            return Decimal("0")
        for balance in accounts:
            if balance.get("currency") == currency:
                try:
                    return Decimal(balance.get("balance", "0"))
                except InvalidOperation:
                    return Decimal("0")
        return Decimal("0")

    def get_order_fill_price(self, order_uuid, retry=7):
        if not order_uuid:
            return None
        for _ in range(retry):
            try:
                order = self.auth_request("GET", "/v1/order", {"uuid": order_uuid})
                if not order:
                    return None
                trades = order.get("trades", [])
                if trades:
                    total_price = sum(Decimal(t["price"]) * Decimal(t["volume"]) for t in trades)
                    total_volume = sum(Decimal(t["volume"]) for t in trades)
                    if total_volume > 0:
                        return total_price / total_volume
            except Exception:
                pass
            time.sleep(1)
        return None


if __name__ == "__main__":
    app = QApplication(sys.argv)
    login_dialog = ApiKeyLoginDialog(
        prefill_access=os.getenv("UPBIT_ACCESS_KEY", "").strip(),
        prefill_secret=os.getenv("UPBIT_SECRET_KEY", "").strip(),
    )
    if login_dialog.exec_() != QDialog.Accepted:
        sys.exit(0)

    access_key, secret_key = login_dialog.get_keys()
    window = UpbitTrader(access_key, secret_key)
    window.show()
    sys.exit(app.exec_())
