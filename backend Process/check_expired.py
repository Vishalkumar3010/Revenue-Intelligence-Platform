import json
import os

def check():
    path = "final_data.json"
    if not os.path.exists(path):
        print("File not found")
        return
    
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    debug_records = data.get("debug_records", [])
    expired = [r for r in debug_records if "expired" in r.get("debug_reason", "").lower()]
    
    print(f"Total debug records: {len(debug_records)}")
    print(f"Expired records: {len(expired)}")
    
    partners = {}
    for r in expired:
        p = r.get("partner_name", "Unknown")
        partners[p] = partners.get(p, 0) + 1
    
    print("\nExpired records by partner:")
    for p, count in partners.items():
        print(f"  - {p}: {count}")

if __name__ == "__main__":
    check()
