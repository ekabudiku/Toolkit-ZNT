import os
import urllib.request
import zipfile
import shutil
import ctypes

def main():
    print("=======================================")
    print("      Toolkit ZNT - Installer/Updater  ")
    print("=======================================\n")

    # Tautan langsung ke repositori Toolkit-ZNT Anda
    VERSION_URL = "https://raw.githubusercontent.com/ekabudiku/Toolkit-ZNT/main/version.txt"
    ZIP_URL = "https://github.com/ekabudiku/Toolkit-ZNT/archive/refs/heads/main.zip"
    
    # Target Folder: Documents/Toolkit ZNT
    doc_path = os.path.expanduser("~\\Documents")
    install_dir = os.path.join(doc_path, "Toolkit ZNT")
    local_version_file = os.path.join(install_dir, "version.txt")

    if not os.path.exists(install_dir):
        os.makedirs(install_dir)
        print("[+] Membangun direktori di: " + install_dir)
        local_ver = "0.0.0"
    else:
        try:
            with open(local_version_file, 'r') as f:
                local_ver = f.read().strip()
        except:
            local_ver = "0.0.0"

    print("Versi Lokal  : " + local_ver)
    print("Mengecek server GitHub...")

    try:
        req = urllib.request.Request(VERSION_URL, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            github_ver = response.read().decode('utf-8').strip()
        print("Versi GitHub : " + github_ver)

        if github_ver != local_ver:
            print("\n[!] Mengunduh pembaruan...")
            zip_path = os.path.join(install_dir, "update.zip")
            
            req = urllib.request.Request(ZIP_URL, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as response, open(zip_path, 'wb') as out_file:
                shutil.copyfileobj(response, out_file)

            print("[+] Mengekstrak file ke direktori lokal...")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                # Bypass folder utama hasil zip dari GitHub
                root_folder = zip_ref.namelist()[0]
                for member in zip_ref.namelist():
                    if member == root_folder: continue
                    
                    target_path = os.path.join(install_dir, member.replace(root_folder, ""))
                    if member.endswith('/'):
                        os.makedirs(target_path, exist_ok=True)
                    else:
                        os.makedirs(os.path.dirname(target_path), exist_ok=True)
                        with zip_ref.open(member) as source, open(target_path, "wb") as target:
                            shutil.copyfileobj(source, target)

            os.remove(zip_path)
            print("\n[V] INSTALASI/UPDATE BERHASIL!")
            ctypes.windll.user32.MessageBoxW(0, f"Toolkit ZNT berhasil diperbarui ke versi {github_ver}!\nLokasi: {install_dir}", "Update Sukses", 0)
        else:
            print("\n[V] Toolkit ZNT Anda sudah dalam versi terbaru.")
            
    except Exception as e:
        print("\n[X] Error: " + str(e))
        print("Pastikan koneksi internet stabil dan file version.txt sudah ada di repositori GitHub.")

    input("\nTekan Enter untuk keluar...")

if __name__ == "__main__":
    main()