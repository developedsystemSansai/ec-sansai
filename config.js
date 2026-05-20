/**
 * config.js — EC Online Submission, Sansai Hospital
 * ─────────────────────────────────────────────────────────────────
 * ไฟล์นี้ถูก gitignore แล้ว — ห้าม commit ขึ้น GitHub
 * ดูตัวอย่างค่าที่ต้องกรอกใน config.example.js
 *
 * วิธีได้ GAS_URL:
 *   Google Apps Script Editor → Deploy → Manage deployments
 *   → เลือก deployment → คัดลอก Web app URL
 *
 * ─────────────────────────────────────────────────────────────────
 * ⚠️  คำเตือน: อย่า hardcode URL นี้ในไฟล์ HTML ใดๆ
 *     ให้แก้ที่นี่ที่เดียว แล้วทุกหน้าจะอ่านค่านี้อัตโนมัติ
 * ─────────────────────────────────────────────────────────────────
 */

window.EC_CONFIG = {
  /**
   * GAS Web App deployment URL
   * รูปแบบ: https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec
   */
  GAS_URL: 'https://script.google.com/macros/s/AKfycby_GFp6PtqRmLWEqayleD9nBscQvBTm288Dm5YMUMSHgXbaie3FTTuvB276hfYN9M69VQ/exec',
};
