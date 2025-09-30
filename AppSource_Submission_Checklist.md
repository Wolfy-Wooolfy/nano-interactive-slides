# ✅ AppSource Submission Checklist

## 1. Partner Center Account
- [ ] إنشاء حساب Partner Center (Individual: 19$ / Company: 99$).  
- [ ] تأكيد الهوية (Identity Verification).  
- [ ] إضافة وسيلة دفع صالحة.  

## 2. Manifest & Add-in Package
- [ ] ملف `manifest.xml` سليم وValidated عبر [Office Add-in Validator](https://aka.ms/officeaddinvalidator).  
- [ ] استخدام **HTTPS** فقط في كل الروابط (no HTTP).  
- [ ] أيقونات add-in (32x32, 64x64, 128x128, 256x256) مضبوطة وبصيغة PNG.  
- [ ] اسم وصفي وواضح (Product Name ≤ 30 حرف).  
- [ ] وصف قصير (≤ 100 حرف) + وصف مطوّل (≤ 4,000 حرف).  
- [ ] تحديد Permissions (Read/Write) بشكل صحيح.  
- [ ] التأكد من عدم وجود API Calls بتفشل أو بتدي Timeout.  

## 3. تجربة المستخدم (UX)
- [ ] الواجهة (Taskpane) تفتح وتشتغل بدون Errors.  
- [ ] الأزرار الأساسية (Start / Stop / Nano Mode) تعمل وظائفها المتوقعة.  
- [ ] في حالة الـ Error → يظهر Message واضح للمستخدم.  
- [ ] تجربة المستخدم بسيطة ومفهومة (No Dead Buttons).  

## 4. السياسات والقوانين
- [ ] Privacy Policy (رابط فعال https://example.com/privacy).  
- [ ] Terms of Use (رابط فعال https://example.com/terms).  
- [ ] Contact Email ظاهر وصالح (support@example.com).  
- [ ] لا يوجد محتوى مسيء / مخالف (صور، نصوص).  

## 5. Publishing Package
- [ ] Screenshots عالية الجودة (3–5 صور) توضح الـ Add-in داخل PowerPoint.  
- [ ] فيديو (اختياري لكن مفيد) يشرح الوظائف الرئيسية.  
- [ ] Keywords مناسبة (مثلاً: AI Slides, Interactive Slides, Productivity).  
- [ ] اختيار الفئة الصحيحة (Productivity → Presentations).  

## 6. اختبار متكامل
- [ ] اختبار الـ Add-in على **PowerPoint Windows** + **PowerPoint Online (Web)**.  
- [ ] التأكد من أن نفس الـ manifest يشتغل على Office 365 (آخر إصدار).  
- [ ] لا توجد Console Errors في DevTools وقت التشغيل.  

---

📌 **نصيحة مهمة:**  
قبل الرفع النهائي، شغّل الأمر:  

```powershell
office-addin-validate manifest.xml
