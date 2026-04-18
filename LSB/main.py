import struct
from PIL import Image

# python main.py input.png payload.bin output.png 0x77 20

def hide_payload_final(image_path, payload_path, output_path, key=0x77, offset_pixels=150):
    with open(payload_path, "rb") as f:
        payload = f.read()
    
    # 1. Шифруем шеллкод
    encrypted_payload = bytearray([b ^ key for b in payload])
    
    # 2. Формируем заголовок: 4 байта (длина шеллкода) + сам шеллкод
    # Длину тоже XOR-им для единообразия
    length_header = struct.pack("<I", len(payload)) # Little-endian 4 bytes
    full_data = bytearray([b ^ key for b in length_header]) + encrypted_payload
    
    # 3. Преобразуем всё в биты
    bits = []
    for byte in full_data:
        for i in range(8):
            bits.append((byte >> i) & 1)

    img = Image.open(image_path).convert("RGB")
    pixels = img.load()
    width, height = img.size
    
    if len(bits) > (width * height - offset_pixels) * 3:
        raise ValueError("Картинка слишком мала!")

    bit_idx = 0
    curr_pix = 0
    for y in range(height):
        for x in range(width):
            if curr_pix < offset_pixels:
                curr_pix += 1
                continue
            
            r, g, b = pixels[x, y]
            channels = [r, g, b]
            for i in range(3):
                if bit_idx < len(bits):
                    channels[i] = (channels[i] & ~1) | bits[bit_idx]
                    bit_idx += 1
            pixels[x, y] = tuple(channels)
            if bit_idx >= len(bits): break
            curr_pix += 1
        if bit_idx >= len(bits): break

    img.save(output_path, "PNG")
    print(f"Записано байт: {len(payload)} (всего с заголовком: {len(full_data)})")

if __name__ == "__main__":
    import sys

    # Expect at least 3 args
    if len(sys.argv) < 4:
        print("Usage: python script.py <image_path> <payload_path> <output_path> [key] [offset]")
        print("Example: python script.py in.png payload.bin out.png 0x77 150")
        sys.exit(1)

    image_path = sys.argv[1]
    payload_path = sys.argv[2]
    output_path = sys.argv[3]

    # Optional args
    key = int(sys.argv[4], 0) if len(sys.argv) > 4 else 0x77
    offset = int(sys.argv[5]) if len(sys.argv) > 5 else 150

    hide_payload_final(
        image_path,
        payload_path,
        output_path,
        key=key,
        offset_pixels=offset
    )