with open("lic_automation_app.py", "r", encoding="utf-8") as f:
    for idx, line in enumerate(f, start=1):
        if "float" in line:
            print(f"Line {idx}: {line.strip()}")
