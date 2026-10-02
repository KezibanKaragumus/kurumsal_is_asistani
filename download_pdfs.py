import os
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin

# Hedef dosyaları kaydetmek için klasör oluştur
download_folder = "belgeler"
os.makedirs(download_folder, exist_ok=True)

# İndirilecek dosya türlerini belirleyin (örn: PDF, DOCX, XLSX vb.)
file_extensions = [".pdf", ".docx", ".xlsx", ".txt", ".jpg", ".png"]

# Çoklu URL listesini belirtin
urls_to_scrape = [
    "https://sayzek.org.tr/"
   
]

downloaded_count = 0

# Sayfaları döngüyle işleme
for main_url in urls_to_scrape:
    print(f"Processing main URL: {main_url}")
    try:
        main_page = requests.get(main_url)
        main_page.raise_for_status()
        soup = BeautifulSoup(main_page.text, "html.parser")
        
        # Sayfadaki tüm linkleri bul
        links = soup.find_all("a", href=True)
        
        for i, link in enumerate(links, start=1):
            file_url = urljoin(main_url, link["href"])
            file_name = link.get_text(strip=True) or f"file_{i}"
            clean_file_name = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in file_name)
            
            # Dosya uzantısını kontrol et
            if any(file_url.lower().endswith(ext) for ext in file_extensions):
                file_extension = os.path.splitext(file_url)[1].lower()
                save_path = os.path.join(download_folder, f"{clean_file_name}{file_extension}")
                
                print(f"Downloading file: {file_url}")
                try:
                    file_response = requests.get(file_url)
                    file_response.raise_for_status()
                    with open(save_path, "wb") as f:
                        f.write(file_response.content)
                    print(f"Saved: {save_path}")
                    downloaded_count += 1
                except Exception as e:
                    print(f"Error downloading file: {e}")
            else:
                # Eğer link bir dosyaya değilse (örneğin alt sayfa), içerikte dosya arayın
                try:
                    sub_page = requests.get(file_url)
                    sub_page.raise_for_status()
                    sub_soup = BeautifulSoup(sub_page.text, "html.parser")
                    
                    for ext in file_extensions:
                        file_links = sub_soup.find_all("a", href=True)
                        for file_link in file_links:
                            sub_file_url = urljoin(file_url, file_link["href"])
                            if sub_file_url.lower().endswith(ext):
                                sub_file_name = file_link.get_text(strip=True) or f"file_{i}"
                                clean_sub_file_name = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in sub_file_name)
                                sub_save_path = os.path.join(download_folder, f"{clean_sub_file_name}{ext}")
                                
                                print(f"Downloading subpage file: {sub_file_url}")
                                try:
                                    sub_file_response = requests.get(sub_file_url)
                                    sub_file_response.raise_for_status()
                                    with open(sub_save_path, "wb") as f:
                                        f.write(sub_file_response.content)
                                    print(f"Saved: {sub_save_path}")
                                    downloaded_count += 1
                                except Exception as e:
                                    print(f"Error downloading subpage file: {e}")
                except Exception as e:
                    print(f"Error accessing subpage: {e}")
    except Exception as e:
        print(f"Error accessing main URL: {e}")

print(f"Download complete. Total files downloaded: {downloaded_count}.")
