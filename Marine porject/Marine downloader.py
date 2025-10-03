import os
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, unquote
import urllib3
from datetime import datetime
import time
import re

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configuration
BASE_URL = "https://msi.admiralty.co.uk"
WEEKLY_URL = "https://msi.admiralty.co.uk/NoticesToMariners/Weekly"
DOWNLOAD_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads", "UK_Marine_Notices")

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-GB,en;q=0.9',
    'Referer': 'https://msi.admiralty.co.uk/NoticesToMariners/Weekly',
}

def create_download_folder():
    """Create folder for downloads"""
    if not os.path.exists(DOWNLOAD_FOLDER):
        os.makedirs(DOWNLOAD_FOLDER)
        print(f"Created folder: {DOWNLOAD_FOLDER}")
    return DOWNLOAD_FOLDER

def extract_pdf_links_from_page(html_content):
    """Extract all PDF download links from the HTML page"""
    soup = BeautifulSoup(html_content, 'html.parser')
    pdf_links = []
    
    # Find all links with specific pattern:
    # <a target="_blank" aria-label="Download File for FILENAME" href="/NoticesToMariners/DownloadFile?...">
    for link in soup.find_all('a', href=True):
        href = link['href']
        aria_label = link.get('aria-label', '')
        
        # Check if this is a download link
        if 'DownloadFile' in href and 'Download' in link.get_text():
            
            # Extract filename from aria-label: "Download File for 41snii25. Link open new tab"
            filename = None
            if 'for' in aria_label.lower():
                # Pattern: "Download File for FILENAME. Link open new tab"
                match = re.search(r'for\s+([a-zA-Z0-9_\-]+)', aria_label)
                if match:
                    filename = match.group(1)
                    # Add .pdf if not present
                    if not filename.endswith('.pdf'):
                        filename += '.pdf'
            
            # Fallback: extract from fileName parameter
            if not filename and 'fileName=' in href:
                match = re.search(r'fileName=([^&]+)', href)
                if match:
                    filename = unquote(match.group(1))
            
            if filename:
                # Build full URL (handle &amp; in HTML)
                href_clean = href.replace('&amp;', '&')
                full_url = urljoin(BASE_URL, href_clean)
                
                pdf_links.append({
                    'filename': filename,
                    'url': full_url
                })
    
    return pdf_links

def download_pdf(session, pdf_info, download_folder):
    """Download a single PDF file"""
    filename = pdf_info['filename']
    url = pdf_info['url']
    filepath = os.path.join(download_folder, filename)
    
    # Skip if already downloaded
    if os.path.exists(filepath):
        file_size = os.path.getsize(filepath) / 1024
        print(f"  SKIP: {filename} (already exists, {file_size:.1f} KB)")
        return True
    
    try:
        print(f"  Downloading: {filename}...", end=' ', flush=True)
        
        response = session.get(url, timeout=60, verify=False, stream=True)
        response.raise_for_status()
        
        # Save file
        with open(filepath, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        
        file_size = os.path.getsize(filepath) / 1024
        print(f"✓ ({file_size:.1f} KB)")
        return True
        
    except Exception as e:
        print(f"✗ Error: {e}")
        # Clean up failed download
        if os.path.exists(filepath):
            os.remove(filepath)
        return False

def main():
    """Main download process"""
    print("=" * 70)
    print("UK Admiralty Weekly Notices - PDF Downloader")
    print("=" * 70)
    
    # Create download folder
    download_folder = create_download_folder()
    print(f"Download location: {download_folder}\n")
    
    # Setup session
    session = requests.Session()
    session.headers.update(HEADERS)
    
    print("Fetching current week's PDFs from Admiralty website...")
    
    try:
        # Get the weekly notices page
        response = session.get(WEEKLY_URL, timeout=30, verify=False)
        response.raise_for_status()
        
        # Save HTML for debugging
        debug_file = os.path.join(download_folder, "page_debug.html")
        with open(debug_file, 'w', encoding='utf-8') as f:
            f.write(response.text)
        print(f"Saved page HTML to: {debug_file}\n")
        
        # Extract PDF links
        pdf_links = extract_pdf_links_from_page(response.text)
        
        if not pdf_links:
            print("❌ No PDF links found on the page.")
            print("\nPossible reasons:")
            print("1. The page uses JavaScript to load content dynamically")
            print("2. The page structure has changed")
            print(f"3. Check the saved HTML file: {debug_file}")
            print("\n💡 Alternative: Download manually from the website:")
            print("   https://msi.admiralty.co.uk/NoticesToMariners/Weekly")
            print("   Click 'Download All' button")
            return
        
        print(f"Found {len(pdf_links)} PDFs:\n")
        for i, pdf in enumerate(pdf_links, 1):
            print(f"  {i}. {pdf['filename']}")
        
        print("\n" + "=" * 70)
        print("Starting downloads...")
        print("=" * 70 + "\n")
        
        # Download each PDF
        success_count = 0
        for i, pdf_info in enumerate(pdf_links, 1):
            print(f"[{i}/{len(pdf_links)}]")
            if download_pdf(session, pdf_info, download_folder):
                success_count += 1
            time.sleep(1)  # Be polite to the server
        
        print("\n" + "=" * 70)
        print(f"Download complete: {success_count}/{len(pdf_links)} successful")
        print(f"Files saved to: {download_folder}")
        print("=" * 70)
        
        # List downloaded files
        if success_count > 0:
            print("\n📁 Downloaded files:")
            for filename in sorted(os.listdir(download_folder)):
                if filename.endswith('.pdf'):
                    filepath = os.path.join(download_folder, filename)
                    size = os.path.getsize(filepath) / 1024
                    print(f"  • {filename} ({size:.1f} KB)")
        
    except requests.exceptions.RequestException as e:
        print(f"❌ Error fetching page: {e}")
        print("\n💡 Try:")
        print("1. Check your internet connection")
        print("2. Visit the website manually: " + WEEKLY_URL)
        print("3. The website might be temporarily down")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()