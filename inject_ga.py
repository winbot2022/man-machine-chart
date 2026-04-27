import os
import pathlib
import streamlit as st

def inject_ga():
    # GA4の測定IDをここに貼り付けてください
    GA_ID = "G-PVGL652YQE" 

    # GA4のタグ（JavaScript）
    GA_JS = f"""
    <script async src="https://www.googletagmanager.com/gtag/js?id={GA_ID}"></script>
    <script>
        window.dataLayer = window.dataLayer || [];
        function gtag(){{dataLayer.push(arguments);}}
        gtag('js', new Date());
        gtag('config', '{GA_ID}');
    </script>
    """

    # Streamlitのインストール先にある index.html を探す
    index_path = pathlib.Path(st.__file__).parent / "static" / "index.html"
    
    with open(index_path, "r") as f:
        html = f.read()

    # すでに書き込み済みでないか確認し、<head> 内に挿入する
    if GA_ID not in html:
        updated_html = html.replace("<head>", "<head>" + GA_JS)
        with open(index_path, "w") as f:
            f.write(updated_html)
        print(f"GA4 ({GA_ID}) を正常に注入しました。")
    else:
        print("GA4 は既に注入済みです。")

if __name__ == "__main__":
    inject_ga()
