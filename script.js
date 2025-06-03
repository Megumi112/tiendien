
let duLieuKH = {};
let lichSuIn = [];
let lichSuThu = [];

function docFileExcel(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet);
        duLieuKH = {};

        const datalist = document.getElementById("danhSachKH");
        datalist.innerHTML = "";

        rows.forEach(row => {
            if (row.MA_KHTT) {
                duLieuKH[row.MA_KHTT] = row;
                const option = document.createElement("option");
                option.value = row.MA_KHTT;
                datalist.appendChild(option);
            }
        });
        alert("Đã tải dữ liệu Excel thành công.");
    };
    reader.readAsArrayBuffer(file);
}

function timKiemMaKH() {
    const ma = document.getElementById("maKH").value.trim();
    if (duLieuKH[ma]) {
        dienThongTin(duLieuKH[ma]);
    } else {
        alert("Không tìm thấy mã khách hàng.");
    }
}

function timKiemTenKH() {
    const ten = document.getElementById("tenKHInput").value.trim().toLowerCase();
    const match = Object.values(duLieuKH).find(kh => kh.TEN_KHTT && kh.TEN_KHTT.toLowerCase().includes(ten));
    if (match) {
        dienThongTin(match);
    } else {
        alert("Không tìm thấy tên khách hàng.");
    }
}

function dienThongTin(data) {
    const ten = data.TEN_KHTT || "";
    const ma = data.MA_KHTT || "";
    const soTien = parseFloat(data.TONG_NOP || 0);
    const phi = 2000;
    const tong = soTien + phi;

    document.getElementById("tenKH").value = ten;
    document.getElementById("soTien").value = soTien.toLocaleString();
    document.getElementById("maKH").value = ma;

    const hoaDon = `BIÊN NHẬN THANH TOÁN TIỀN ĐIỆN\n\nMã KH: ${ma}\nTên KH: ${ten}\nSố Tiền: ${soTien.toLocaleString()} ₫\nPhí: ${phi.toLocaleString()} ₫\nTổng Tiền: ${tong.toLocaleString()} ₫\n\nĐÃ THANH TOÁN\nNgày ${new Date().toLocaleDateString("vi-VN")}`;
    document.getElementById("hoaDon").innerText = hoaDon;
}

function inHoaDon() {
    const hoaDon = document.getElementById("hoaDon").innerText;
    if (hoaDon.trim()) {
        lichSuIn.push(hoaDon);
        capNhatLichSu();
        let wnd = window.open('', '_blank');
        wnd.document.write('<pre>' + hoaDon + '</pre>');
        wnd.print();
        wnd.close();
    } else {
        alert("Không có hóa đơn để in.");
    }
}

function xacNhanThu() {
    const hoaDon = document.getElementById("hoaDon").innerText;
    if (hoaDon.trim()) {
        lichSuThu.push(hoaDon);
        capNhatLichSu();
    } else {
        alert("Không có hóa đơn để xác nhận.");
    }
}

function capNhatLichSu() {
    document.getElementById("lichSuIn").value = lichSuIn.join("\n---\n");
    document.getElementById("lichSuThu").value = lichSuThu.join("\n---\n");
}

function xuatLichSu(loai) {
    const noiDung = loai === 'in' ? lichSuIn : lichSuThu;
    const blob = new Blob([noiDung.join("\n---\n")], { type: "text/plain;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = loai === 'in' ? "lich_su_in.txt" : "lich_su_thu.txt";
    a.click();
}

function xuatLichSuExcel() {
    const wb = XLSX.utils.book_new();
    const wsIn = XLSX.utils.json_to_sheet(lichSuIn.map(chuyenChuoiThanhObject));
    const wsThu = XLSX.utils.json_to_sheet(lichSuThu.map(chuyenChuoiThanhObject));
    XLSX.utils.book_append_sheet(wb, wsIn, "Lịch sử in");
    XLSX.utils.book_append_sheet(wb, wsThu, "Lịch sử thu");
    XLSX.writeFile(wb, "lich_su_in_thu.xlsx");
}

function chuyenChuoiThanhObject(hoaDon) {
    const obj = {};
    const lines = hoaDon.split('\n');
    lines.forEach(line => {
        if (line.includes("Mã KH:")) obj["Mã KH"] = line.split("Mã KH:")[1].trim();
        if (line.includes("Tên KH:")) obj["Tên KH"] = line.split("Tên KH:")[1].trim();
        if (line.includes("Số Tiền:")) obj["Số tiền"] = line.split("Số Tiền:")[1].replace("₫", "").trim();
        if (line.includes("Phí:")) obj["Phí"] = line.split("Phí:")[1].replace("₫", "").trim();
        if (line.includes("Tổng Tiền:")) obj["Tổng"] = line.split("Tổng Tiền:")[1].replace("₫", "").trim();
        if (line.includes("Ngày")) obj["Ngày"] = line.replace("Ngày ", "").trim();
    });
    return obj;
}
