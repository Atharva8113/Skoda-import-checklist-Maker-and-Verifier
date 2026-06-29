def safe_float(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip()
    if not val_str:
        return 0.0
    # Remove commas and spaces
    val_str = val_str.replace(',', '').replace(' ', '')
    try:
        return float(val_str)
    except ValueError:
        return 0.0

# Test values from user screenshot and edge cases
test_cases = [
    ' 10,81,920 ',
    ' 3,451,545.67 ',
    '10,00,000.00',
    '   ',
    None,
    12345.67,
    ' -12,345 ',
]

for tc in test_cases:
    res = safe_float(tc)
    print(f"Input: {repr(tc)} => Output: {res} (Type: {type(res).__name__})")
