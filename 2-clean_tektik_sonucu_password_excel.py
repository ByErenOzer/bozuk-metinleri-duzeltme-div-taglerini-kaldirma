import re, html
import pandas as pd
from openpyxl import load_workbook
import msoffcrypto
import io

def clean_text(s):
    if s is None:
        return s
    if not isinstance(s, str):
        s = str(s)
    t = s
    for _ in range(2):
        t2 = html.unescape(t)
        if t2 == t:
            break
        t = t2
    t = re.sub(r'(?i)<br\s*/?>', '\n', t)
    t = re.sub(r'(?i)</\s*(div|p|span)\s*>', '\n', t)
    t = re.sub(r'(?i)<\s*(div|p|span)[^>]*>', '', t)
    t = re.sub(r'(?i)<[^>]+>', '', t)
    t = t.replace('\u00a0', ' ')
    t = re.sub(r'\s+\n', '\n', t)
    t = re.sub(r'\n\s+', '\n', t)
    t = re.sub(r'\n{3,}', '\n\n', t)
    t = re.sub(r'[ \t]{2,}', ' ', t)
    t = t.strip()
    return t

def process_workbook(src, dst):
    # Decrypt the password-protected Excel file
    decrypted_file = io.BytesIO()
    with open(src, 'rb') as f:
        file = msoffcrypto.OfficeFile(f)
        file.load_key(password='sb.271906')
        file.decrypt(decrypted_file)
    decrypted_file.seek(0)
    
    # Read all sheets from decrypted Excel file
    excel_file = pd.ExcelFile(decrypted_file, engine='openpyxl')
    changed = {}
    
    with pd.ExcelWriter(dst, engine='openpyxl') as writer:
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            count = 0
            
            # Find target columns and clean them
            for col in df.columns:
                if col and str(col).strip().lower() in ('tetkik_sonucu','tetkit_sonucu'):
                    # Clean the column values
                    for idx, value in df[col].items():
                        before = value
                        after = clean_text(before)
                        df.at[idx, col] = after
                        if before != after:
                            count += 1
                    # Rename column
                    df = df.rename(columns={col: 'tetkik_sonucu_temiz'})
            
            # Save the processed sheet
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            if count > 0:
                changed[sheet_name] = count
    
    print('saved', dst)
    for k,v in changed.items():
        print(f'{k}: {v}')
    return changed

if __name__ == '__main__':
    src = r'c:\\Users\\doganeren.ozer\\Desktop\\div_clean_text\\patolojiLLMguncel.xlsx'
    dst = r'c:\\Users\\doganeren.ozer\\Desktop\\div_clean_text\\patolojiLLMguncel-temiz.xlsx'
    process_workbook(src, dst)