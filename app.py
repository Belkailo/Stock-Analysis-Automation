import yfinance as yf
import pandas as pd
import numpy as np

# სტოკების სია
stock_symbols = ["AMD", "PLTR", "NVDA", "SMCI", "ANET","ACLX","AGYS","AISP","TSLA","META","AMZN","MSFT","SMTC"]

# ცხრილის მონაცემები
data = []

# სტოკების მონაცემების გამოყვანა
for symbol in stock_symbols:
    try:
        stock = yf.Ticker(symbol)
        stock_data = stock.history(period="1y")  # 1 წლის მონაცემები

        # Moving Averages (50 და 200 დღის)
        stock_data['MA50'] = stock_data['Close'].rolling(window=50).mean()
        stock_data['MA200'] = stock_data['Close'].rolling(window=200).mean()

        # RSI (14 დღიანი)
        delta = stock_data['Close'].diff()
        gain = delta.where(delta > 0, 0)
        loss = -delta.where(delta < 0, 0)

        avg_gain = gain.rolling(window=14).mean()
        avg_loss = loss.rolling(window=14).mean()

        rs = avg_gain / avg_loss
        stock_data['RSI'] = 100 - (100 / (1 + rs))

        # MACD
        stock_data['EMA12'] = stock_data['Close'].ewm(span=12, adjust=False).mean()
        stock_data['EMA26'] = stock_data['Close'].ewm(span=26, adjust=False).mean()
        stock_data['MACD'] = stock_data['EMA12'] - stock_data['EMA26']
        stock_data['Signal_Line'] = stock_data['MACD'].ewm(span=9, adjust=False).mean()

        # Bollinger Bands
        stock_data['Middle_Band'] = stock_data['Close'].rolling(window=20).mean()
        stock_data['Upper_Band'] = stock_data['Middle_Band'] + 2 * stock_data['Close'].rolling(window=20).std()
        stock_data['Lower_Band'] = stock_data['Middle_Band'] - 2 * stock_data['Close'].rolling(window=20).std()

        # ATR (Average True Range)
        stock_data['High-Low'] = stock_data['High'] - stock_data['Low']
        stock_data['High-Close'] = abs(stock_data['High'] - stock_data['Close'].shift())
        stock_data['Low-Close'] = abs(stock_data['Low'] - stock_data['Close'].shift())
        stock_data['True Range'] = stock_data[['High-Low', 'High-Close', 'Low-Close']].max(axis=1)
        stock_data['ATR'] = stock_data['True Range'].rolling(window=14).mean()

        # პროცენტული ცვლილება (1 თვე)
        stock_data['Monthly_Change (%)'] = stock_data['Close'].pct_change(periods=30) * 100

        # ბოლო მონაცემები
        current_price = stock_data['Close'].iloc[-1]
        ma50 = stock_data['MA50'].iloc[-1]
        ma200 = stock_data['MA200'].iloc[-1]
        rsi = stock_data['RSI'].iloc[-1]
        macd = stock_data['MACD'].iloc[-1]
        signal_line = stock_data['Signal_Line'].iloc[-1]
        upper_band = stock_data['Upper_Band'].iloc[-1]
        lower_band = stock_data['Lower_Band'].iloc[-1]
        atr = stock_data['ATR'].iloc[-1]
        monthly_change = stock_data['Monthly_Change (%)'].iloc[-1]

        # სიგნალები
        signal = ""
        if ma50 > ma200:
            signal += "MA50 crosses above MA200: Buy signal\n"
        elif ma50 < ma200:
            signal += "MA50 crosses below MA200: Sell signal\n"

        if rsi > 70:
            signal += "RSI above 70: Overbought (Sell signal)\n"
        elif rsi < 30:
            signal += "RSI below 30: Oversold (Buy signal)\n"

        if macd > signal_line:
            signal += "MACD above Signal Line: Buy signal\n"
        elif macd < signal_line:
            signal += "MACD below Signal Line: Sell signal\n"

        # მონაცემები ცხრილში
        data.append({
            "Symbol": symbol,
            "Current Price ($)": current_price,
            "MA50": ma50,
            "MA200": ma200,
            "RSI (14)": rsi,
            "MACD": macd,
            "Signal_Line": signal_line,
            "Upper Band": upper_band,
            "Lower Band": lower_band,
            "ATR": atr,
            "Monthly Change (%)": monthly_change,
            "Signal": signal.strip()
        })

    except Exception as e:
        print(f"Error processing {symbol}: {e}")

# ცხრილის შექმნა
df = pd.DataFrame(data)
df.to_excel('stock.xlsx')

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ელ. ფოსტის პარამეტრები
sender_email = " "
receiver_email = " "
app_password = " "  # თქვენი "App Password"

# შეტყობინების შექმნა
message = MIMEMultipart()
message['From'] = sender_email
message['To'] = receiver_email
message['Subject'] = "Stock Technical Analysis Results"

# ტექსტური მონაცემები (მაგალითად, ცხრილის ანალიზი)
body = df.to_html()  # ცხრილის HTML ფორმატში გადაქცევა
message.attach(MIMEText(body, 'html'))

# SMTP კავშირი და ელ. ფოსტის გაგზავნა
try:
    server = smtplib.SMTP_SSL('smtp.mail.yahoo.com', 465)  # SSL კავშირი
    server.login(sender_email, app_password)  # თქვენი "App Password"-ით ავთენთიფიკაცია
    text = message.as_string()
    server.sendmail(sender_email, receiver_email, text)
    server.quit()
    print("Email sent successfully!")
except Exception as e:
    print(f"Failed to send email: {e}")
