---
name: Lead_Scoring_Skill
description: Bộ quy tắc nghiệp vụ và hướng dẫn kỹ thuật xây dựng hệ thống AI tự động chấm điểm khách hàng tiềm năng (Lead Scoring) bất động sản, tích hợp Web App duyệt dữ liệu (Human-in-the-loop) và xuất Excel.
---

# HƯỚNG DẪN XÂY DỰNG HỆ THỐNG AI LEAD SCORING & AUTOMATION

Tài liệu này chuẩn hóa quy trình nghiệp vụ, các câu lệnh gợi ý (Prompts) cho AI, thiết kế cơ sở dữ liệu, kiến trúc Web App Human-in-the-loop, và hướng dẫn xuất báo cáo Excel cho ngành Bất động sản.

---

## MỤC LỤC
1. [Nghiệp vụ Chấm điểm Khách hàng (Business Rules)](#1-nghiep-vu-cham-diem-khach-hang-business-rules)
2. [Thiết kế Prompt AI Lead Scoring](#2-thiet-ke-prompt-ai-lead-scoring)
3. [Quy trình Tự động hóa Lấy Dữ liệu (Google Sheets Integration)](#3-quy-trinh-tu-dong-hoa-lay-du-lieu-google-sheets-integration)
4. [Kiến trúc Web App Human-in-the-loop (Flask + UI Glassmorphism)](#4-kien-truc-web-app-human-in-the-loop-flask--ui-glassmorphism)
5. [Tích hợp Xuất Excel Chuyên nghiệp (Formatting Output)](#5-tich-hop-xuat-excel-chuyen-nghiep-formatting-output)

---

## 1. NGHIỆP VỤ CHẤM ĐIỂM KHÁCH HÀNG (BUSINESS RULES)

Quy tắc chấm điểm dựa trên thông tin mô tả nhu cầu của khách hàng. Điểm mặc định ban đầu là **50 điểm**.

### 1.1. Tiêu chí Cộng 50 Điểm (VIP/Siêu Tiềm Năng)
Cộng 50 điểm nếu khách hàng thuộc một hoặc nhiều nhóm sau:
*   **Ngân sách lớn:** Có đề cập số tiền cụ thể từ 20 tỷ đồng trở lên hoặc các cụm từ "tài chính mạnh", "không thành vấn đề".
*   **Loại hình cao cấp:** Tìm kiếm "Biệt thự đơn lập", "Penthouse", "Shophouse mặt đường lớn", "Quỹ đất công nghiệp", "Sàn văn phòng diện tích lớn".
*   **Vị trí đắc địa:** Yêu cầu các khu vực VIP như "Quận 1", "Ven sông", "Vinhomes Ocean Park", "Phú Mỹ Hưng".
*   **Đối tượng khách hàng:** Đề cập là "Chủ doanh nghiệp", "Nhà đầu tư chuyên nghiệp", "Mua sỉ", "Mua số lượng lớn".
*   **Tính cấp thiết & Minh bạch:** Yêu cầu "Pháp lý chuẩn 100%", "Sổ hồng riêng", "Muốn gặp trực tiếp chủ đầu tư để đàm phán".

### 1.2. Tiêu chí Trừ 50 Điểm (Khách hàng Rác/Không Tiềm Năng)
Trừ 50 điểm nếu khách hàng thuộc một hoặc nhiều nhóm sau:
*   **Yêu cầu phi thực tế:** Tìm mua bất động sản với giá thấp vô lý so với thị trường (Ví dụ: Nhà Quận 1 giá 1-2 tỷ, nhà trung tâm có sân vườn hồ bơi giá vài trăm triệu).
*   **Không có nhu cầu:** "Nhầm số", "Không có nhu cầu", "Dữ liệu cũ", "Nhầm ngành".
*   **Khách hàng không thiện chí:** "Hỏi giá cho vui", "Chưa có ý định mua", "Thái độ không hợp tác".
*   **Spam/Quảng cáo:** Nội dung chứa các dịch vụ khác như "Bảo hiểm", "Vay vốn", "Mời chào dịch vụ".
*   **Thông tin liên lạc lỗi:** "Thuê bao", "Gọi nhiều lần không bắt máy", "Không phản hồi Zalo".

### 1.3. Các Trường hợp Khác (Giữ Nguyên Điểm Hoặc Cộng Ít: 50-60 Điểm)
*   Khách hàng tìm mua chung cư, nhà phố tầm trung (3-10 tỷ).
*   Khách hàng cần vay ngân hàng, đang cân nhắc chính sách.
*   Khách hàng có nhu cầu thực nhưng cần tư vấn thêm về pháp lý hoặc vị trí.

---

## 2. THIẾT KẾ PROMPT AI LEAD SCORING

Để AI đánh giá một cách khách quan, chúng ta sử dụng kỹ thuật Structuring Output định dạng JSON để dễ dàng tích hợp vào hệ thống backend.

### 2.1. System Prompt

```markdown
Bạn là chuyên gia tư vấn Bất động sản và chuyên viên Phân tích dữ liệu khách hàng. 
Nhiệm vụ của bạn là đánh giá và chấm điểm mức độ tiềm năng của khách hàng (Lead) dựa trên nội dung mô tả nhu cầu chi tiết.

### Quy tắc chấm điểm (Bắt đầu từ mức cơ sở là 50 điểm):
1. CỘNG 50 ĐIỂM (VIP) nếu phát hiện:
   - Ngân sách >= 20 tỷ hoặc "tài chính mạnh", "không thành vấn đề".
   - Loại hình: "Biệt thự đơn lập", "Penthouse", "Shophouse mặt đường lớn", "Quỹ đất công nghiệp", "Sàn văn phòng diện tích lớn".
   - Vị trí: "Quận 1", "Ven sông", "Vinhomes Ocean Park", "Phú Mỹ Hưng".
   - Đối tượng: "Chủ doanh nghiệp", "Nhà đầu tư chuyên nghiệp", "Mua sỉ", "Mua số lượng lớn".
   - Tính cấp thiết: "Pháp lý chuẩn 100%", "Sổ hồng riêng", "Muốn gặp trực tiếp chủ đầu tư để đàm phán".

2. TRỪ 50 ĐIỂM (RÁC/KHÔNG TIỀM NĂNG) nếu phát hiện:
   - Phi thực tế: Nhà Quận 1 giá 1-2 tỷ, nhà trung tâm có sân vườn hồ bơi giá vài trăm triệu,...
   - Không có nhu cầu: "Nhầm số", "Không có nhu cầu", "Dữ liệu cũ", "Nhầm ngành".
   - Thiếu thiện chí: "Hỏi giá cho vui", "Chưa có ý định mua", "Thái độ không hợp tác".
   - Spam/Quảng cáo: Chứa nội dung về "Bảo hiểm", "Vay vốn", "Mời chào dịch vụ".
   - Lỗi liên lạc: "Thuê bao", "Gọi nhiều lần không bắt máy", "Không phản hồi Zalo".

3. CÁC TRƯỜNG HỢP KHÁC (GIỮ NGUYÊN ĐIỂM):
   - Mua chung cư, nhà phố tầm trung 3-10 tỷ.
   - Cần vay ngân hàng, cân nhắc chính sách.
   - Nhu cầu thực nhưng cần tư vấn thêm.
   Điểm số cuối cùng phải giới hạn trong khoảng từ 0 đến 100 điểm.

### Định dạng đầu ra:
Bạn bắt buộc phải trả về chuỗi JSON duy nhất, không kèm markdown code block (không dùng ```json), có cấu trúc sau:
{
  "score": <số nguyên từ 0 đến 100>,
  "classification": <"VIP" | "Tiềm năng" | "Rác">,
  "reason": "<Giải thích chi tiết lý do cộng/trừ điểm dựa trên bộ quy tắc>",
  "extracted_info": {
    "budget": "<Thông tin ngân sách trích xuất>",
    "property_type": "<Loại hình bất động sản>",
    "location": "<Vị trí yêu cầu>",
    "urgency": "<Mức độ cấp thiết>",
    "is_spam_or_invalid": <true/false>
  }
}
```

### 2.2. User Prompt Template

```markdown
Đánh giá khách hàng có thông tin sau:
- Tên khách hàng: {customer_name}
- Nhu cầu/Ghi chú: {customer_note}
```

---

## 3. QUY TRÌNH TỰ ĐỘNG HÓA LẤY DỮ LIỆU (GOOGLE SHEETS INTEGRATION)

Dữ liệu nguồn được quản lý trên Google Sheets. Ta phát triển script Python để kết nối và tải dữ liệu.

### 3.1. Phương pháp 1: Đọc công khai thông qua export CSV (Khuyên dùng vì đơn giản)
Nếu link Google Sheets được chia sẻ ở quyền "Bất kỳ ai có liên kết đều có thể xem":
```python
import pandas as pd

def fetch_data_from_google_sheet(sheet_url):
    # Chuyển đổi link edit thành link export CSV
    csv_url = sheet_url.replace('/edit?usp=sharing', '/export?format=csv')
    csv_url = csv_url.replace('/edit?gid=0#gid=0', '/export?format=csv&gid=0')
    if '/edit' in csv_url and '/export' not in csv_url:
        csv_url = csv_url.split('/edit')[0] + '/export?format=csv'
        
    df = pd.read_csv(csv_url)
    return df
```

### 3.2. Phương pháp 2: Sử dụng API `gspread` (Khi sheet ở chế độ riêng tư)
Yêu cầu cài đặt thư viện `gspread oauth2client` và tệp cấu hình `credentials.json`.
```python
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def fetch_private_sheet(sheet_key, credentials_path):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_key).sheet1
    return sheet.get_all_records()
```

---

## 4. KIẾN TRÚC WEB APP HUMAN-IN-THE-LOOP

Giao diện Web App cho phép người quản lý kiểm duyệt và chỉnh sửa thủ công điểm số hoặc phân loại mà AI đưa ra trước khi kết xuất file cuối cùng.

### 4.1. Cơ sở dữ liệu tạm thời (In-memory / JSON file)
Để đơn giản hóa việc lưu trữ tạm thời trong buổi học, dữ liệu sau khi chấm điểm từ AI được lưu vào một danh sách trong bộ nhớ (In-memory storage) hoặc tệp `leads_data.json` để đồng bộ trạng thái khi người dùng chỉnh sửa trên giao diện.

### 4.2. Triển khai API Backend (Flask)
```python
from flask import Flask, jsonify, request, render_template
import pandas as pd

app = Flask(__name__)
leads_database = []  # Lưu dữ liệu khách hàng trong phiên chạy

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/leads', methods=['GET'])
def get_leads():
    return jsonify(leads_database)

@app.route('/api/leads/<int:lead_id>', methods=['PUT'])
def update_lead(lead_id):
    # Human-in-the-loop: Cập nhật trạng thái/điểm từ người dùng
    data = request.json
    for lead in leads_database:
        if lead['id'] == lead_id:
            lead['score'] = int(data.get('score', lead['score']))
            lead['classification'] = data.get('classification', lead['classification'])
            lead['human_reviewed'] = True
            return jsonify({"status": "success", "lead": lead})
    return jsonify({"status": "error", "message": "Lead not found"}), 404
```

### 4.3. UI Glassmorphism & Dark Mode
Giao diện bảng điều khiển cần áp dụng các thuộc tính CSS Glassmorphism:
```css
.card-glass {
    background: rgba(15, 23, 42, 0.45);
    backdrop-filter: blur(16px);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 12px;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
}
```

---

## 5. TÍCH HỢP XUẤT EXCEL CHUYÊN NGHIỆP

Dữ liệu bàn giao cho bộ phận Sales cần được định dạng đẹp mắt trên Excel bằng thư viện `openpyxl`.

```python
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

def export_leads_to_formatted_excel(leads, file_path):
    df = pd.DataFrame(leads)
    
    # Tạo Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Khach Hang tiem Nang"
    
    # Grid lines visible
    ws.views.sheetView[0].showGridLines = True
    
    # Header styling
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Navy Accent
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # Ghi dữ liệu từ DataFrame
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
        
    # Định dạng các cột và dòng
    thin_border = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD')
    )
    
    for col in ws.columns:
        # Tự động căn chỉnh độ rộng cột
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
    for cell in ws[1]: # Dòng tiêu đề
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        
    # Lưu file
    wb.save(file_path)
```
