import docx

doc = docx.Document(r"c:\Users\mksin\Desktop\AI hardened BMS\Cyber_Hardened_BMS_Manual.docx")

print("Listing all main chapters/headings:")
for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if any(txt.startswith(prefix) for prefix in ["Chapter", "APPENDIX", "Appendix", "TABLE OF CONTENTS"]):
        # Safe print replacing non-ascii
        safe_txt = txt.encode("ascii", "replace").decode("ascii")
        print(f"P{i}: {safe_txt}")
