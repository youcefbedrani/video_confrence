import hashlib
import random
import string

# IMPORTANT: This must match the salt in installer.py
LICENSE_SALT = "Nextcloud_DZ_Secret_2026"

def generate_checksum(key_data):
    combined = key_data + LICENSE_SALT
    return hashlib.sha256(combined.encode()).hexdigest()[:10].upper()

def generate_key(type_code):
    """
    types: 
    TR = Trial (15 days)
    AN = Annual (1 year)
    LT = Lifetime
    """
    # 1. Create a random part
    random_part = ''.join(random.choices(string.ascii_uppercase + string.digits, k=5))
    
    # 2. Build the base key
    base_key = f"NC-{type_code.upper()}-{random_part}"
    
    # 3. Generate checksum
    checksum = generate_checksum(base_key)
    
    # 4. Return full key
    return f"{base_key}-{checksum}"

if __name__ == "__main__":
    print("-" * 30)
    print(" Nextcloud Key Generator")
    print("-" * 30)
    print("1. Trial (15 day)")
    print("2. Annual (1 year)")
    print("3. Lifetime")
    
    choice = input("\nSelect key type (1-3): ")
    type_map = {"1": "TR", "2": "AN", "3": "LT"}
    
    if choice in type_map:
        t = type_map[choice]
        count = int(input("How many keys to generate? ") or 1)
        print("\nGenerated Keys (Copy ONLY the part starting with NC-):")
        for _ in range(count):
            print(f"[{t}] {generate_key(t)}")
    else:
        print("Invalid choice.")
    print("\nKeep this script safe — don't share it with clients!")
