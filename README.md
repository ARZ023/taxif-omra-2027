# سحب العمرة 2027 — TaxiF

موقع ثابت (HTML/CSS/JS) لسحب جائزة العمرة 2027 للكباتن المرخّصين والمحرّرين.
لا يحتاج سيرفر ولا بناء — والملفات تُقرأ داخل المتصفح فقط ولا تُرفع لأي مكان.

## الاستخدام
1. افتح الصفحة، ثم اسحب ملف **JSON** وملف **Excel** وأفلتهما على العجلة.
2. راجع «شروط الأهلية» (يُضبط تلقائياً على أعمدة مرخّص/محرّر) وعدّل القيم المسموحة إن لزم.
3. اضغط **ابدأ السحب** (أو مفتاح المسافة).
4. «تصدير Excel» يحفظ قائمة الفائزين مع كل بياناتهم.

## النشر على GitHub Pages
```bash
git init && git add . && git commit -m "TaxiF Omra draw"
git branch -M main
git remote add origin https://github.com/<USER>/<REPO>.git
git push -u origin main
```
ثم من GitHub: **Settings → Pages → Deploy from a branch → main / (root)**.
الرابط سيكون `https://<USER>.github.io/<REPO>/`.

> ملفات الكباتن (xlsx / json / csv) مستثناة في `.gitignore` حتى لا تُرفع بالخطأ.
