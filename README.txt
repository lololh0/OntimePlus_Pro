Feedback AI Project — Strong UI Version
====================================

Run:
1) pip install -r requirements.txt
2) python -m uvicorn feedback_backend:app --reload
3) In another PowerShell: python -m http.server 5500
4) Open:
   http://127.0.0.1:5500/index-en.html
   http://127.0.0.1:5500/index-ar.html

Improvements:
- Compact Anonymous Comments UI in Arabic and English.
- Active search, department filter, and sentiment filter.
- Comment cards in responsive grid.
- Sentiment colors and smoother hover animation.
- Auto-refresh for KPI/comments when backend is running.
