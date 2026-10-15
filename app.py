import streamlit as st
import pandas as pd
from yt_dlp import YoutubeDL
import io

st.set_page_config(page_title="Multi YouTube Comment Scraper", page_icon="📊", layout="centered")

st.title("📊 Multi-URL YouTube Comment Scraper to Excel")
st.write("ஒன்றுக்கு மேற்பட்ட YouTube வீடியோ லிங்குகளை ஒரே நேரத்தில் கொடுத்து கமெண்ட்ஸ்களை Excel-ல் டவுன்லோட் செய்யுங்கள்.")

# Multiple URLs input using text_area
urls_input = st.text_area(
    "YouTube Video URLs (ஒவ்வொரு வரியிலும் ஒரு லிங்க் வீதம் போடவும்):",
    placeholder="https://www.youtube.com/watch?v=...\nhttps://www.youtube.com/watch?v=..."
)

max_comments = st.slider("ஒவ்வொரு வீடியோவிற்கும் எத்தனை கமெண்ட்கள் வேண்டும்?", min_value=10, max_value=1000, value=50, step=10)

def scrape_multiple_videos(urls, limit):
    all_comments = []
    
    # Split input text by lines and clean empty spaces
    url_list = [url.strip() for url in urls.split("\n") if url.strip()]
    
    ydl_opts = {
        'getcomments': True,
        'skip_download': True,
        'extract_flat': True,
    }
    
    for url in url_list:
        try:
            with YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(url, download=False)
                video_title = info.get('title', 'YouTube Video')
                comments = info.get('comments', [])
                
                for c in comments[:limit]:
                    all_comments.append({
                        'Video URL': url,
                        'Video Title': video_title,
                        'Author': c.get('author', 'Anonymous'),
                        'Comment': c.get('text', ''),
                        'Likes': c.get('like_count', 0),
                        'Published At': c.get('timestamp', '')
                    })
        except Exception as e:
            st.warning(f"இந்த லிங்கில் பிழை ஏற்பட்டது ({url}): {str(e)}")
            
    return pd.DataFrame(all_comments)

if st.button("Extract All Comments"):
    if not urls_input:
        st.warning("தயவுசெய்து குறைந்தது ஒரு YouTube URL-ஐயாவது உள்ளிடவும்!")
    else:
        with st.spinner("அனைத்து வீடியோ கமெண்ட்களும் சேகரிக்கப்படுகின்றன... சற்று காத்திருக்கவும்..."):
            df = scrape_multiple_videos(urls_input, max_comments)
            
            if df.empty:
                st.error("கமெண்ட்கள் எதுவும் கிடைக்கவில்லை அல்லது லிங்குகள் தவறானவை.")
            else:
                st.success(f"வெற்றிகரமாக மொத்தம் {len(df)} கமெண்ட்கள் சேகரிக்கப்பட்டன!")
                st.dataframe(df.head()) # Preview table
                
                # Convert DataFrame to Excel in memory
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='All_Comments')
                
                # Download Button
                st.download_button(
                    label="📥 அனைத்து கமெண்ட்ஸ்களையும் Excel-ல் பதிவிறக்க",
                    data=buffer.getvalue(),
                    file_name="multiple_youtube_comments.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
