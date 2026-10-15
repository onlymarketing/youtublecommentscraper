import streamlit as st
import pandas as pd
from yt_dlp import YoutubeDL
import io

# Page Config
st.set_page_config(page_title="YouTube Comment Scraper", page_icon="📊", layout="centered")

st.title("📊 YouTube Comment Scraper to Excel")
st.write("Commercial Vehicle Insurance & YouTube வீடியோக்களின் கமெண்ட்ஸ்களை சுலபமாக Excel-ல் டவுன்லோட் செய்யுங்கள்.")

# User Input URL
url = st.text_input("YouTube Video URL-ஐ உள்ளிடவும்:", placeholder="https://www.youtube.com/watch?v=...")

# Max comments count
max_comments = st.slider("எத்தனை கமெண்ட்கள் சேகரிக்க வேண்டும்?", min_value=10, max_value=2000, value=100, step=10)

def scrape_youtube_comments(video_url, limit):
    ydl_opts = {
        'getcomments': True,
        'skip_download': True,
        'extract_flat': True,
    }
    
    comments_list = []
    
    with YoutubeDL(ydl_opts) as ydl:
        info = ydl.extract_info(video_url, download=False)
        video_title = info.get('title', 'YouTube Video')
        comments = info.get('comments', [])
        
        for c in comments[:limit]:
            comments_list.append({
                'Video Title': video_title,
                'Author': c.get('author', 'Anonymous'),
                'Comment': c.get('text', ''),
                'Likes': c.get('like_count', 0),
                'Published At': c.get('timestamp', '')
            })
            
    return video_title, pd.DataFrame(comments_list)

if st.button("Scrape Comments"):
    if not url:
        st.warning("தயவுசெய்து ஒரு செல்லுபடியாகும் YouTube URL-ஐ உள்ளிடவும்!")
    else:
        with st.spinner("கமெண்ட்கள் சேகரிக்கப்படுகின்றன... சற்று காத்திருக்கவும்..."):
            try:
                video_title, df = scrape_youtube_comments(url, max_comments)
                
                if df.empty:
                    st.error("இந்த வீடியோவில் கமெண்ட்கள் எதுவும் கிடைக்கவில்லை அல்லது கமெண்ட்கள் முடக்கப்பட்டுள்ளன.")
                else:
                    st.success(f"வெற்றிகரமாக {len(df)} கமெண்ட்கள் சேகரிக்கப்பட்டன!")
                    st.dataframe(df.head()) # Preview table
                    
                    # Convert DataFrame to Excel in memory
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='Comments')
                    
                    # Download Button
                    st.download_button(
                        label="📥 Excel கோப்பாக பதிவிறக்க (Download Excel)",
                        data=buffer.getvalue(),
                        file_name="youtube_comments.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"பிழை ஏற்பட்டது: {str(e)}")
