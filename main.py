# pip install streamlit python-docx

import streamlit as st
from docx import Document
from io import BytesIO
# import pythoncom

def markdown_to_word(markdown_text):
    # Wordドキュメントを作成
    doc = Document()
    
    # マークダウンの段落を処理
    for line in markdown_text.split('\n'):
        # 空行は無視
        if line.strip() == '':
            continue
            
        # 見出しの処理 (簡易的に#の数で判断)
        if line.startswith('#'):
            level = len(line.split(' ')[0])
            heading_text = line.replace('#', '').strip()
            if level == 1:
                doc.add_heading(heading_text, level=0)
            elif level == 2:
                doc.add_heading(heading_text, level=1)
            elif level >= 3:
                doc.add_heading(heading_text, level=2)
        else:
            # 通常の段落
            doc.add_paragraph(line)
    
    return doc

def save_docx(doc):
    # Wordドキュメントをバイナリデータとして保存
    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# StreamlitアプリのUI
st.title('マークダウンからWordへ変換')
st.write('マークダウン形式のテキストを入力し、Wordファイルとしてダウンロードできます。')

# テキスト入力エリア
markdown_text = st.text_area(
    'マークダウンテキストを入力してください',
    height=300,
    value='# 見出し1\n\n## 見出し2\n\nこれは通常の段落です。\n\n- リスト項目1\n- リスト項目2'
)

# 変換ボタン
if st.button('Wordファイルに変換'):
    if markdown_text:
        try:
            # pythoncom.CoInitialize()  # COMの初期化 (Windows環境で必要になる場合があります)
            doc = markdown_to_word(markdown_text)
            word_file = save_docx(doc)
            
            st.success('変換が完了しました！')
            st.download_button(
                label='Wordファイルをダウンロード',
                data=word_file,
                file_name='output.docx',
                mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        except Exception as e:
            st.error(f'エラーが発生しました: {e}')
    else:
        st.warning('マークダウンテキストを入力してください')