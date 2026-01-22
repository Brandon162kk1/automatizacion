import os

# Rutas donde buscar supervisord.conf
folders = ["./Facturas", "./Cuotas"]

for folder in folders:
    conf_path = os.path.join(folder, "supervisord.conf")
    if os.path.exists(conf_path):
        with open(conf_path, "rb") as f:
            content = f.read()
        # Detectar BOM UTF-8 (EF BB BF)
        if content.startswith(b"\xef\xbb\xbf"):
            print(f"🔧 Limpiando BOM en: {conf_path}")
            content = content[3:]  # quitar BOM
            with open(conf_path, "wb") as f:
                f.write(content)
        else:
            print(f"✅ Sin BOM: {conf_path}")
    else:
        print(f"⚠️ No existe: {conf_path}")
