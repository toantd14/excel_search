const xlsx = require('xlsx');
const _ = require('lodash');
const moment = require('moment');

// Hàm chuyển đổi số serial Excel thành ngày
function convertExcelDate(excelSerial) {
  const excelStartDate = new Date(1899, 11, 30); // Ngày bắt đầu của Excel (30/12/1899)
  const convertedDate = new Date(excelStartDate.getTime() + (excelSerial - 1) * 86400000); // Mỗi số đại diện cho một ngày
  return convertedDate;
}

// Đọc dữ liệu từ file Excel
const workbook = xlsx.readFile('data/sao_ke_bao_so_3_10_09_2024.xlsx'); // Thay đường dẫn đúng với file Excel của bạn
const sheetName = workbook.SheetNames[0]; // Chọn sheet đầu tiên
const sheet = workbook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(sheet);

// Lấy từ khóa và số lượng bản ghi từ dòng lệnh
const maxRecords = parseInt(process.argv[2]); // Số bản ghi tối đa lấy ra
const searchString = process.argv[3];         // Từ khóa tìm kiếm

// Hàm tìm kiếm theo từ khóa và in ngay khi tìm thấy
function searchAndPrint(keyword, limit) {
  let count = 0;

  // Duyệt qua từng dòng và tìm kiếm
  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    // Kiểm tra nếu bất kỳ cột nào chứa từ khóa
    if (_.some(row, value => typeof value === 'string' && value.toLowerCase().includes(keyword.toLowerCase()))) {
      count++;

      // Xử lý định dạng ngày nếu là số serial
      let formattedDate;
      if (typeof row.date === 'number') {
        const dateObj = convertExcelDate(row.date); // Chuyển số serial thành đối tượng Date
        formattedDate = moment(dateObj).format('DD/MM/YYYY'); // Định dạng ngày thành DD/MM/YYYY
      } else {
        formattedDate = row.date || 'N/A';
      }

      // Xử lý số tiền (đảm bảo đúng định dạng số)
      const formattedAmount = row.amount ? parseFloat(row.amount).toLocaleString('vi-VN') + ' VND' : 'N/A';

      // Kiểm tra và in ra các trường hợp trống
      const transactionCode = row.code || 'N/A';
      const transactionNotes = row.notes || 'Không có diễn giải';

      // In ra thông tin theo thứ tự đã yêu cầu
      console.log(`Transaction ${count}:`);
      console.log('---------------------------------------');
      console.log(`Mã giao dịch: ${transactionCode}`);
      console.log(`Ngày giao dịch: ${formattedDate}`);
      console.log(`Số tiền giao dịch: ${formattedAmount}`);
      console.log(`Diễn giải giao dịch: ${transactionNotes}`);
      console.log(); // Tạo dòng trống giữa các giao dịch

      // Dừng tìm kiếm nếu đã đạt đến số lượng bản ghi yêu cầu
      if (count >= limit) {
        return; // Dừng hàm ngay khi đủ bản ghi
      }
    }
  }

  // Nếu không tìm thấy bản ghi nào
  if (count === 0) {
    console.log(`Không tìm thấy bản ghi nào với từ khóa: ${keyword}`);
  }
}

// Tìm kiếm và in kết
searchAndPrint(searchString, maxRecords);
