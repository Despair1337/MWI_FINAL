import struct
from PIL import Image

# python test.py output.png payload.bin 0x77 20

def extract_payload(image_path, key=0x77, offset_pixels=150):
    img = Image.open(image_path).convert("RGB")
    pixels = img.load()
    width, height = img.size

    bits = []
    curr_pix = 0

    # 1. Extract all LSB bits after offset
    for y in range(height):
        for x in range(width):
            if curr_pix < offset_pixels:
                curr_pix += 1
                continue

            r, g, b = pixels[x, y]
            for channel in (r, g, b):
                bits.append(channel & 1)

            curr_pix += 1

    # 2. Convert bits → bytes
    data = bytearray()
    for i in range(0, len(bits), 8):
        byte = 0
        for j in range(8):
            if i + j < len(bits):
                byte |= (bits[i + j] << j)
        data.append(byte)

    # 3. Decode length header (first 4 bytes, XORed)
    header_enc = data[:4]
    header = bytearray([b ^ key for b in header_enc])
    payload_len = struct.unpack("<I", header)[0]

    # 4. Extract payload (also XORed)
    payload_enc = data[4:4 + payload_len]
    payload = bytearray([b ^ key for b in payload_enc])

    return payload, payload_len


def verify_payload(image_path, expected_payload_path, key=0x77, offset_pixels=150):
    extracted_payload, payload_len = extract_payload(
        image_path, key, offset_pixels
    )

    with open(expected_payload_path, "rb") as f:
        expected = f.read()

    print(f"Extracted length: {payload_len}")
    print(f"Expected length:  {len(expected)}")

    if extracted_payload == expected:
        print("✅ Payload matches!")
    else:
        print("❌ Payload does NOT match!")

        # Optional: show first mismatch
        for i in range(min(len(extracted_payload), len(expected))):
            if extracted_payload[i] != expected[i]:
                print(f"Mismatch at byte {i}: {extracted_payload[i]} != {expected[i]}")
                break


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("Usage: python verify.py <image_path> <payload_path> [key] [offset]")
        sys.exit(1)

    image_path = sys.argv[1]
    payload_path = sys.argv[2]

    key = int(sys.argv[3], 0) if len(sys.argv) > 3 else 0x77
    offset = int(sys.argv[4]) if len(sys.argv) > 4 else 150

    verify_payload(image_path, payload_path, key, offset)