const CONFIG = {
    // URL ของ Google Apps Script Web App
    GAS_URL: 'https://script.google.com/macros/s/AKfycbxaZu6t_xT89qbGV612-Q8Ng_d-FqJ5V3QAoAwqsT5k9cg4O5iK_OlTTtcTXXgAzQXz/exec',

    // URL ของ Local Hardware Service (Python)
    LOCAL_API: 'http://127.0.0.1:5000',

    // ตั้งค่าแผนที่เริ่มต้น
    MAP_CENTER: [17.221, 102.427],
    MAP_ZOOM: 13,

    // ระบบสีหลัก
    COLORS: {
        PRIMARY: '#0d6efd',
        SUCCESS: '#198754',
        DANGER: '#dc3545',
        WARNING: '#ffc107',
        DARK: '#212529'
    }
};

// ฟังก์ชันช่วยเหลือสำหรับพิมพ์ Logs แบบสวยงาม
const AppLog = (msg, data) => {
    console.log(`%c[LandLease System]%c ${msg}`, 'color: #0d6efd; font-weight: bold', 'color: inherit', data || '');
};
