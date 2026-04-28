/**
 * Upload food images to Google Drive via Apps Script
 * Usage: node tools/upload_images.js
 * 
 * ขั้นตอน:
 * 1. เปิด POS /dev → ไปหน้าตั้งค่า → คลิก "📁 สร้างโฟลเดอร์รูปอาหาร"
 * 2. จะได้ลิงก์โฟลเดอร์ → เปิดใน Google Drive
 * 3. อัปโหลดรูปที่ต้องการเข้าไปในโฟลเดอร์
 * 4. ตั้งชื่อไฟล์ตาม product_id เช่น P001.png, P002.png
 * 5. กลับมาที่ POS → คลิก "🔗 อัปเดตลิงก์รูป"
 */

const fs = require('fs');
const path = require('path');

// Map of product IDs to local image files
const imageMap = {
  // เพิ่ม product_id: 'path_to_image' ตรงนี้
  // เช่น: 'P001': 'C:/path/to/laab_moo.png'
};

// Generated images directory
const imgDir = 'C:\\Users\\Pratueang\\.gemini\\antigravity\\brain\\ea1bb62b-51b5-423b-95a8-590050851324';

// List available generated images
console.log('📁 รูปอาหารที่ generate แล้ว:');
try {
  const files = fs.readdirSync(imgDir).filter(f => f.endsWith('.png') && !f.startsWith('media_'));
  files.forEach(f => console.log('  - ' + f));
  console.log('\n💡 วิธีใช้:');
  console.log('1. เปิด Google Drive → โฟลเดอร์ POS_FoodImages');
  console.log('2. อัปโหลดรูปจากโฟลเดอร์นี้:');
  console.log('   ' + imgDir);
  console.log('3. ตั้งชื่อไฟล์ตาม product_id (P001.png, P002.png ฯลฯ)');
} catch(e) {
  console.log('ไม่พบโฟลเดอร์รูป');
}
