# বিদ্যালয় ম্যানেজমেন্ট সিস্টেম (Vidyalaya)

একটি সম্পূর্ণ স্কুল ম্যানেজমেন্ট অ্যাপ — Dashboard, Students, Teachers, Finance, Due, Master Fee, Committee, Reports সবকিছু একসাথে। ডেটা সেভ হয় আপনার নিজের **Google Sheet**-এ (ফ্রি ডেটাবেজ)।

## ফাইল
- **index.html** — পুরো অ্যাপটাই এই একটা ফাইলে। শুধু এটা ব্রাউজারে খুললেই চলবে।
- **AppsScript-Code.gs** — Google Sheet-কে ডেটাবেজ বানানোর কোড (Google-এ পেস্ট করতে হবে)।

আগে এটা দুইটা ফাইলে (index.html + app.js) ভাগ করা ছিল, যার কারণে সরাসরি ডাবল-ক্লিক করে খুললে **সাদা/খালি পেইজ** দেখাচ্ছিল — ব্রাউজার নিরাপত্তার কারণে local ফাইল থেকে ফাইল লোড আটকে দেয়। এখন সব একসাথে **index.html**-এ বসিয়ে দেওয়া হয়েছে, তাই এখন থেকে সরাসরি ওপেন করলেই কাজ করবে।

---

## সহজ সেটআপ (৩ ধাপ)

### ধাপ ১: Google Sheet ডেটাবেজ তৈরি করুন
1. [sheets.google.com](https://sheets.google.com)-এ গিয়ে একটা নতুন খালি Sheet খুলুন।
2. উপরে **Extensions → Apps Script**-এ ক্লিক করুন।
3. যা লেখা আছে সব মুছে, **AppsScript-Code.gs** ফাইলের পুরো কোডটা পেস্ট করে দিন।
4. **Deploy → New deployment** ক্লিক করুন:
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. **Deploy** চাপুন, পারমিশন দিন, তারপর যে URL-টা দেখাবে (শেষে `/exec` থাকবে) সেটা কপি করুন।

### ধাপ ২: URL-টা index.html-এ বসান
1. **index.html** ফাইলটা যেকোনো Text Editor (Notepad, VS Code ইত্যাদি) দিয়ে খুলুন।
2. `Ctrl+F` চেপে `PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE` লেখাটা খুঁজুন।
3. এই লাইনটা পাবেন:
   ```js
   const DEFAULT_SHEETS_URL = "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";
   ```
4. `PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE` এর জায়গায় আপনার কপি করা URL-টা বসান, যেমন:
   ```js
   const DEFAULT_SHEETS_URL = "https://script.google.com/macros/s/AKfycb.../exec";
   ```
5. ফাইলটা সেভ করুন। ব্যাস, এখন **index.html** ডাবল-ক্লিক করে খুললেই ডেটা সরাসরি আপনার Google Sheet-এ সেভ হবে।

> চাইলে স্কুলের নামও একই জায়গায় বসাতে পারেন (`DEFAULT_SCHOOL_NAME` লাইনে) — এটা ID কার্ড, ইনভয়েসে দেখাবে।

### ধাপ ৩: ছবি/ডকুমেন্ট আপলোড (Token লাগবে না)
আগের ভার্সনে GitHub Token লাগত — এখন সেটা বাদ দেওয়া হয়েছে। এখন দুইভাবে ছবি যোগ করতে পারবেন:

**অপশন A — GitHub-এ নিজে আপলোড করে লিংক পেস্ট করুন:**
1. আপনার GitHub রিপোতে (github.com-এ গিয়ে) ছবি/ফাইলটা ম্যানুয়ালি আপলোড করুন (drag & drop)।
2. আপলোড হওয়া ফাইলে ক্লিক করে **Raw** বাটনে ক্লিক করুন — যে লিংকটা পাবেন সেটা কপি করুন।
3. অ্যাপের ছবি ফিল্ডে (Student/Teacher ফর্মে) সেই লিংকটা পেস্ট করে দিন।

**অপশন B — সরাসরি Embed করুন (GitHub-ই লাগবে না):**
1. ফর্মে **"Or embed a file"** বাটনে ক্লিক করুন।
2. আপনার কম্পিউটার থেকে ছবিটা সিলেক্ট করুন — এটা সরাসরি সেভ হয়ে যাবে, কোনো লিংকের দরকার নেই।

---

## ব্যবহার
- বাম পাশের মেনু থেকে Dashboard, Students, Teachers, Finance, Due, Master Fee, Committee, Reports — সব মডিউলে যেতে পারবেন।
- যেকোনো ID Card, Invoice, Admit Card, বা Letter-এ **Print / Save as PDF** বাটনে ক্লিক করলে সেটা প্রিন্ট বা PDF হিসেবে সেভ করতে পারবেন।
- **Settings** পেইজ থেকেও চাইলে পরে URL পরিবর্তন করতে পারবেন (ব্রাউজারে সেভ থাকবে)।

## নোট
- যদি Google Sheet URL না বসান, অ্যাপ **Offline Mode**-এ চলবে (ডেটা শুধু ওই মুহূর্তে থাকবে, পেইজ রিফ্রেশ করলে হারিয়ে যাবে)।
- সবাইকে এক জায়গা থেকে ব্যবহার করাতে চাইলে, এই **index.html** ফাইলটা GitHub Pages বা যেকোনো ফ্রি হোস্টিং-এ আপলোড করে দিতে পারেন — সবার ডেটা একই Google Sheet-এ জমা হবে।
