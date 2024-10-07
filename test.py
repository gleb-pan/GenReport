import uuid

def get_mac_address():
    mac = uuid.getnode()
    # Convert the MAC address to a human-readable format
    mac_address = ':'.join(f'{(mac >> i) & 0xff:02x}' for i in range(0, 48, 8))
    return mac_address

if __name__ == '__main__':
    print(get_mac_address())