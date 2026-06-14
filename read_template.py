import zipfile, re

path = r"d:\MCP\Claude Code\mau bieu giam dinh pti\assets\CV-Ngân hàng.docx"
with zipfile.ZipFile(path, 'r') as z:
    with z.open('word/document.xml') as f:
        content = f.read().decode('utf-8')

# Extract all placeholders
placeholders = re.findall(r'\{[^}]+\}', content)
print("PLACEHOLDERS:", sorted(set(placeholders)))

# Extract plain text
text = re.sub(r'<[^>]+>', '', content)
text = re.sub(r'\s+', ' ', text)
print("\nTEXT (first 3000):", text[:3000])
