import pandas as pd
import time
import requests
import os
from tkinter import Tk, filedialog
from datetime import datetime

# プロキシー設定（社内用）
proxies = {
    "http": "http://proxy-apj.kyndryl.net:8080",
    "https": "http://proxy-apj.kyndryl.net:8080"
}

# API Key（組織APIキーに置き換え）
api_key = "xxxxxxxxxxxxxxxxx"

# リクエストヘッダー（Bearer認証を使用）
headers = {
    "Accept": "application/json",
    "Content-Type": "application/json",
    "Authorization": f"Bearer {api_key}"
}

# 組織ID（固定値）
org_id = "xxxxxxxxxxxx"

# ファイル選択ダイアログ
Tk().withdraw()
file_path = filedialog.askopenfilename(title="Select the CSV file for suspension", filetypes=[("CSV files", "*.csv")])

if not file_path:
    print("No file selected. Exiting process.")
    exit()

# CSV読み込み
df_input = pd.read_csv(file_path)

# 出力ファイル名（同じフォルダに保存）
source_dir = os.path.dirname(file_path)
result_file = os.path.join(source_dir, "UserSuspend_result.csv")

results = []

# ユーザーごとに suspend 実行
for index, row in df_input.iterrows():
    account_id = row["accountId"]
    email = row["email"]
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    url = f"https://api.atlassian.com/admin/v1/orgs/{org_id}/directory/users/{account_id}/suspend-access"

    try:
        response = requests.post(url, headers=headers, proxies=proxies)

        if response.status_code == 200:
            msg = f"{now}-{email} suspended, status code {response.status_code}"
        elif response.status_code == 400:
            msg = f"{now}-{email} Bad request, status code {response.status_code}"
        elif response.status_code == 401:
            msg = f"{now}-{email} unauthorized, status code {response.status_code}"
        elif response.status_code == 403:
            msg = f"{now}-{email} Forbidden, status code {response.status_code}"
        elif response.status_code == 404:
            msg = f"{now}-{email} Not Found, status code {response.status_code}"
        elif response.status_code == 429:
            msg = f"{now}-Too many API requests, status code {response.status_code}"
        elif response.status_code == 500:
            msg = f"{now}-Internal error, status code {response.status_code}"
        else:
            msg = f"{now}-{email} suspend failed - HTTP {response.status_code}: {response.text}"

    except Exception as e:
        msg = f"{now}-{email} suspend error: {str(e)}"

    print(msg)
    results.append({
        "accountId": account_id,
        "email": email,
        "action": "suspend",
        "result": msg
    })

    time.sleep(1)

# 結果を保存（追記モード、ヘッダー付き/無しを自動判断）
if results:
    df_result = pd.DataFrame(results)
    write_header = not os.path.exists(result_file)
    df_result.to_csv(result_file, mode="a", header=write_header, index=False)
    print(f"{len(results)} entries appended to {result_file}.")
else:
    print("No suspension attempted.")

input("処理が完了しました。画面を閉じてください.")
