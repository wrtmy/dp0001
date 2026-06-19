import pdfplumber
import re
import sys

def extract_invoice_info(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()
        
        if not text:
            print("未提取到文本")
            return

        print("=" * 70)
        print(f"📄 文件: {pdf_path}")
        print("=" * 70)
        
        full_text = ' '.join(text.split())
        full_text = full_text.replace('价 税 合 计', '价税合计')
        full_text = full_text.replace('（ 小 写 ）', '（小写）')
        
        # ---- 1. 发票号码 ----
        invoice_no = ""
        match = re.search(r'发票号码[：:]\s*(\d+)', full_text)
        if match:
            invoice_no = match.group(1)
        
        # ---- 2. 开票日期 ----
        invoice_date = ""
        match = re.search(r'开票日期[：:]\s*(\d{4}年\d{2}月\d{2}日)', full_text)
        if match:
            invoice_date = match.group(1)
        
        # ---- 3. 购方名称 ----
        buyer = ""
        match = re.search(r'购\s*名称[：:]\s*([^\s]+(?:有限公司|公司|厂|店|部))', full_text)
        if not match:
            match = re.search(r'买\s*名称[：:]\s*([^\s]+(?:有限公司|公司|厂|店|部))', full_text)
        if match:
            buyer = match.group(1)
        
        # ---- 4. 销方名称 ----
        seller = ""
        match = re.search(r'销\s*名称[：:]\s*([^\s]+(?:有限公司|公司|厂|店|部))', full_text)
        if not match:
            match = re.search(r'售\s*名称[：:]\s*([^\s]+(?:有限公司|公司|厂|店|部))', full_text)
        if match:
            seller = match.group(1)
        
        # ---- 5. 金额 ----
        amount = ""
        match = re.search(r'合\s*计\s*[¥￥]?\s*([\d.]+)', full_text)
        if match:
            amount = match.group(1)
        else:
            match = re.search(r'合计\s*[¥￥]?\s*([\d.]+)', full_text)
            if match:
                amount = match.group(1)
        
        # ---- 6. 税额 ----
        tax = ""
        match = re.search(r'合\s*计\s*[¥￥]?\s*[\d.]+\s*[¥￥]?\s*([\d.]+)', full_text)
        if match:
            tax = match.group(1)
        else:
            match = re.search(r'税额[：:]\s*[¥￥]?\s*([\d.]+)', full_text)
            if match:
                tax = match.group(1)
        
        # ---- 7. 价税合计 ----
        total_amount = ""
        all_numbers = re.findall(r'[¥￥]\s*([\d.]+)', full_text)
        if all_numbers:
            total_amount = all_numbers[-1]
        
        # ---- 输出结果 ----
        print("🔍 【提取结果】")
        print("=" * 70)
        print(f"发票号码: {invoice_no}")
        print(f"开票日期: {invoice_date}")
        print(f"购方名称: {buyer}")
        print(f"销方名称: {seller}")
        print(f"金额: {amount}")
        print(f"税额: {tax}")
        print(f"价税合计: {total_amount}")
        print("=" * 70)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python3 inspector.py <PDF文件路径>")
        print("示例: python3 inspector.py cz.pdf")
        sys.exit(1)
    
    pdf_file = sys.argv[1]
    extract_invoice_info(pdf_file)