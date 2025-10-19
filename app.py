from flask import Flask, request, render_template
import msoffcrypto
import io
import pandas as pd
import os

app = Flask(__name__)

PASSWORD = "m3@9B$#*52K&692v"

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return "❌ 没有文件被上传", 400

    file = request.files['file']
    filename = file.filename

    if not (filename.lower().endswith('.xlsx') or filename.lower().endswith('.xls')):
        return "❌ 请上传 Excel 文件 (.xlsx 或 .xls)", 400

    file_type = request.form.get('filetype')
    if file_type not in ['type1', 'type2']:
        return "❌ 文件类型未选择或无效", 400

    try:
        if file_type == 'type1':
            office_file = msoffcrypto.OfficeFile(file.stream)
            office_file.load_key(password=PASSWORD)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            df = pd.read_excel(decrypted)

            carton_col = df.columns[2]
            code_col = df.columns[5]

            df_unique = df.drop_duplicates(subset=carton_col)
            counts = df_unique[code_col].value_counts().sort_index()

            city_map = {
                850: "WEST VALLEY", 855: "WEST VALLEY",
                940: "SAN FRANCISCO", 949: "SAN FRANCISCO",
                829: "SALT LAKE CITY", 840: "SALT LAKE CITY",
                920: "SAN DIEGO",
                890: "LAS VEGAS",
                932: "BAKERSFIELD",
                980: "WA", 982: "WA", 983: "WA",
                970: "OR"
            }

            result_lines = []

            for code, count in counts.items():
                try:
                    code_int = int(code)
                except:
                    continue
                city = city_map.get(code_int)
                if city:
                    result_lines.append(f"{city} {code_int}-{count}")

            result_lines.sort()

            return f"""
            <div style="text-align: center; font-family: monospace; white-space: pre-line;">
            {'\n'.join(result_lines)}
            </div>
            """

        elif file_type == 'type2':
            df = pd.read_excel(file.stream, skiprows=9)
            df.iloc[:, 1] = df.iloc[:, 1].fillna(method='ffill')

            carton_col = df.columns[1]
            tail_col = df.columns[18]

            df[tail_col] = df[tail_col].astype(str).str.strip().str.upper()

            ags_boxes = 0
            usps_boxes = 0

            for carton, group in df.groupby(carton_col):
                unique_tails = set(group[tail_col].dropna().unique())
                if len(unique_tails) == 1:
                    tail = list(unique_tails)[0]
                    if 'AGS' in tail:
                        ags_boxes += 1
                    elif 'USPS' in tail:
                        usps_boxes += 1
                else:
                    print(f"⚠️ 大箱 {carton} 出现混合尾端：{unique_tails}（忽略）")

            result_lines = []
            if ags_boxes > 0:
                result_lines.append(f"BBC SPX-{ags_boxes}")
            if usps_boxes > 0:
                result_lines.append(f"BBC USPS-{usps_boxes}")

            return f"""
            <div style="text-align: center; font-family: monospace; white-space: pre-line;">
            {'\n'.join(result_lines)}
            </div>
            """

    except Exception as e:
        return f"❌ 处理错误：{str(e)}", 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
