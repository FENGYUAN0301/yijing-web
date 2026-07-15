#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
易经算网站 - 本地服务器启动脚本
支持历史事件Excel缓存 + 64卦解释Excel调用
"""
import http.server
import socketserver
import webbrowser
import os
import threading
import time
import json
import pandas as pd
from urllib.parse import parse_qs, urlparse

PORT = 8000
DIRECTORY = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(DIRECTORY, "history_cache.xlsx")
GUA64_FILE = os.path.join(DIRECTORY, "gua64.xlsx")

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_POST(self):
        if self.path == '/api/save-history':
            try:
                content_length = int(self.headers.get('Content-Length', 0))
                post_data = self.rfile.read(content_length).decode('utf-8')
                data = json.loads(post_data)

                month = data.get('month')
                day = data.get('day')
                events = data.get('events', [])
                source = data.get('source', 'unknown')

                if not month or not day or not events:
                    self.send_error(400, "Missing data")
                    return

                if os.path.exists(CACHE_FILE):
                    df = pd.read_excel(CACHE_FILE)
                else:
                    df = pd.DataFrame(columns=['month', 'day', 'year', 'text', 'source'])

                df = df[(df['month'] != month) | (df['day'] != day)]

                new_rows = []
                for e in events:
                    new_rows.append({
                        'month': month,
                        'day': day,
                        'year': e.get('year', ''),
                        'text': e.get('text', ''),
                        'source': source
                    })

                df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
                df.to_excel(CACHE_FILE, index=False)

                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"success": True, "saved": len(new_rows)}).encode())
                print(f"[历史缓存] 已保存 {month}月{day}日 {len(new_rows)} 条事件 (来源:{source})")

            except Exception as e:
                self.send_error(500, str(e))
        else:
            self.send_error(404)

    def do_GET(self):
        parsed = urlparse(self.path)

        if parsed.path == '/api/history-cache':
            try:
                params = parse_qs(parsed.query)
                month = int(params.get('month', ['0'])[0])
                day = int(params.get('day', ['0'])[0])

                if not os.path.exists(CACHE_FILE):
                    self.send_response(404)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({"found": False, "reason": "no_cache_file"}).encode())
                    return

                df = pd.read_excel(CACHE_FILE)
                df_day = df[(df['month'] == month) & (df['day'] == day)]

                if len(df_day) == 0:
                    self.send_response(404)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({"found": False, "reason": "no_data"}).encode())
                    return

                events = []
                for _, row in df_day.iterrows():
                    events.append({
                        "year": int(row['year']) if pd.notna(row['year']) else 0,
                        "text": str(row['text'])
                    })

                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({
                    "found": True,
                    "month": month,
                    "day": day,
                    "events": events,
                    "count": len(events)
                }, ensure_ascii=False).encode())
                print(f"[历史缓存] 读取 {month}月{day}日 {len(events)} 条事件")

            except Exception as e:
                self.send_error(500, str(e))

        elif parsed.path == '/api/gua64':
            try:
                params = parse_qs(parsed.query)
                gua_name = params.get('name', [''])[0]

                if not os.path.exists(GUA64_FILE):
                    self.send_response(404)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({"found": False, "reason": "no_gua64_file"}).encode())
                    return

                df = pd.read_excel(GUA64_FILE)

                if gua_name:
                    df_gua = df[df['卦名'] == gua_name]
                else:
                    # 返回所有卦名列表
                    gua_list = df['卦名'].tolist()
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({
                        "found": True,
                        "count": len(gua_list),
                        "gua_list": gua_list
                    }, ensure_ascii=False).encode())
                    return

                if len(df_gua) == 0:
                    self.send_response(404)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({"found": False, "reason": "gua_not_found"}).encode())
                    return

                row = df_gua.iloc[0]
                result = {
                    "found": True,
                    "卦名": str(row['卦名']),
                    "上卦": str(row['上卦']),
                    "下卦": str(row['下卦']),
                    "卦象": str(row['卦象']),
                    "总体运势": str(row['总体运势']),
                    "学业": str(row['学业']),
                    "官运": str(row['官运']),
                    "财运": str(row['财运']),
                    "姻缘": str(row['姻缘']),
                    "寻人失物": str(row['寻人失物']),
                    "健康": str(row['健康']),
                    "事业": str(row['事业']),
                    "出行": str(row['出行'])
                }

                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps(result, ensure_ascii=False).encode())
                print(f"[64卦] 读取卦象: {gua_name}")

            except Exception as e:
                self.send_error(500, str(e))

        else:
            super().do_GET()

def start_server():
    # 初始化缓存文件
    if not os.path.exists(CACHE_FILE):
        df = pd.DataFrame(columns=['month', 'day', 'year', 'text', 'source'])
        df.to_excel(CACHE_FILE, index=False)
        print(f"[初始化] 已创建历史事件缓存文件: {CACHE_FILE}")
    else:
        df = pd.read_excel(CACHE_FILE)
        unique_days = df.groupby(['month', 'day']).size().shape[0] if len(df) > 0 else 0
        print(f"[初始化] 已加载历史事件缓存，共 {len(df)} 条记录，{unique_days} 个日期")

    # 检查64卦文件
    if os.path.exists(GUA64_FILE):
        df_gua = pd.read_excel(GUA64_FILE)
        print(f"[初始化] 已加载64卦解释，共 {len(df_gua)} 卦")
    else:
        print(f"[初始化] 未找到64卦解释文件: {GUA64_FILE}")

    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        print(f"\n服务器已启动: http://localhost:{PORT}")
        print(f"目录: {DIRECTORY}")
        print(f"历史缓存: {CACHE_FILE}")
        print(f"64卦解释: {GUA64_FILE}")
        print("按 Ctrl+C 停止服务器\n")
        httpd.serve_forever()

if __name__ == "__main__":
    print("=" * 50)
    print("     易经算网站 - 本地服务器")
    print("     支持历史事件缓存 + 64卦解释")
    print("=" * 50)

    server_thread = threading.Thread(target=start_server, daemon=True)
    server_thread.start()
    time.sleep(1)

    url = f"http://localhost:{PORT}"
    print(f"正在打开浏览器: {url}")
    webbrowser.open(url)

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n服务器已停止")
