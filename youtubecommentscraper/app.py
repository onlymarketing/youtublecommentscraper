import streamlit as st
import pandas as pd
from yt_dlp import YoutubeDL
import io

st.set_page_config(page_title="YouTube Comment Scraper", page_icon="📊")

st.title("📊 YouTube Comment Scraper to Excel")
st.write("Commercial Vehicle Insurance & YouTube வீடியோக்களின் கமெண்ட்ஸ்களை சுலபமாக Excel-ல் டவுன்லோட் செய்யுங்கள்.")

url = st.text_input("YouTube Video URL-ஐ போடவும்:")
max_comments = st.slider("எத்தனை கமெண்ட்கள் வேண்டும்?", min_value=10, max_value=2000, value=100, step=10)

def scrape_youtube_comments(video_url, limit):
    ydl_opts = {'getcomments': True, 'skip_download': True, 'extract_flat': True}
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
    return pd.DataFrame(comments_list)

if st.button("Extract Comments"):
    if url:
        with st.spinner("கமெண்ட்கள் சேகரிக்கப்படுகின்றன..."):
            try:
                df = scrape_youtube_comments(url, max_comments)
                if not df.empty:
                    st.success(f"{len(df)} கமெண்ட்கள் வெற்றிகரமாக சேகரிக்கப்பட்டன!")
                    
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='Comments')
                    
                    st.download_button(
                        label="📥 Download Excel (.xlsx)",
                        data=buffer.getvalue(),
                        file_name="youtube_comments.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("கமெண்ட்கள் எதுவும் கிடைக்கவில்லை.")
            except Exception as e:
                st.error(f"பிழை: {str(e)}")
    else:
        st.warning("தயவுசெய்து ஒரு URL-ஐ உள்ளிடவும்!")