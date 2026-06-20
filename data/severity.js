/*
 * Bảng tiêu chuẩn đánh giá mức độ nghiêm trọng (S)
 * Trích từ: GL SQS0811 - Tiêu chuẩn đánh giá mức độ nghiêm trọng
 *
 * Mỗi mục:
 *   rank     : cấp độ S (1..10)
 *   category : nhóm "Ảnh hưởng" (cột trái của bảng)
 *   text     : nội dung tiêu chuẩn (dùng cho ý 2 của cột "Ảnh hưởng" trong P-FMEA)
 *   scope    : 'product' = ảnh hưởng đến sản phẩm (khách hàng)
 *              'process' = ảnh hưởng đến công đoạn (chế tạo/lắp ráp)
 */
window.SEVERITY_TABLE = [
  // ---- Ảnh hưởng đến SẢN PHẨM (khách hàng) ----
  { scope: 'product', rank: 10, category: 'Không thể thỏa mãn hạng mục yêu cầu của an toàn và/hoặc quy định',
    text: 'Dạng hỏng hóc mang tính tiềm ẩn gây ảnh hưởng đến thao tác an toàn của xe, và/hoặc không phù hợp với quy định của chính phủ mà không có dấu hiệu báo trước' },
  { scope: 'product', rank: 9, category: 'Không thể thỏa mãn hạng mục yêu cầu của an toàn và/hoặc quy định',
    text: 'Dạng hỏng hóc mang tính tiềm ẩn gây ảnh hưởng đến thao tác an toàn của xe, và/hoặc không phù hợp với quy định của chính phủ có dấu hiệu báo trước' },
  { scope: 'product', rank: 8, category: 'Làm mất hoặc làm giảm chức năng chính',
    text: 'Mất chức năng chính (không thể thao tác xe được, không ảnh hưởng đến thao tác an toàn của xe)' },
  { scope: 'product', rank: 7, category: 'Làm mất hoặc làm giảm chức năng chính',
    text: 'Làm giảm chức năng chính (có thể thao tác xe, tuy nhiên, mức độ tính năng suy giảm)' },
  { scope: 'product', rank: 6, category: 'Làm mất hoặc làm giảm chức năng thứ 2',
    text: 'Làm mất chức năng thứ 2 (có thể thao tác xe, tuy nhiên, chức năng liên quan đến tính thoải mái, tính tiện lợi không hoạt động)' },
  { scope: 'product', rank: 5, category: 'Làm mất hoặc làm giảm chức năng thứ 2',
    text: 'Suy giảm chức năng thứ 2 (có thể thao tác xe, tuy nhiên, mức độ tính năng liên quan đến tính thoải mái, tính tiện lợi bị giảm sút)' },
  { scope: 'product', rank: 4, category: 'Khó chịu',
    text: 'Có lỗi mặt ngoài hoặc tiếng kêu, có thể thao tác xe được, có điểm không phù hợp trên sản phẩm mà hầu hết khách hàng (hơn 75%) nhận ra' },
  { scope: 'product', rank: 3, category: 'Khó chịu',
    text: 'Có lỗi mặt ngoài hoặc tiếng kêu, có thể thao tác xe được, có điểm không phù hợp trên sản phẩm mà nhiều khách hàng (50%) nhận ra' },
  { scope: 'product', rank: 2, category: 'Khó chịu',
    text: 'Có lỗi mặt ngoài hoặc tiếng kêu, có thể thao tác xe được, có điểm không phù hợp trên sản phẩm mà những khách hàng có khả năng phân biệt nội dung đó (dưới 25%) mới nhận ra' },
  { scope: 'product', rank: 1, category: 'Không có ảnh hưởng',
    text: 'Không có ảnh hưởng mà có thể nhận thấy được' },

  // ---- Ảnh hưởng đến CÔNG ĐOẠN (chế tạo/lắp ráp) ----
  { scope: 'process', rank: 10, category: 'Không thể thỏa mãn hạng mục yêu cầu của an toàn và/hoặc quy định',
    text: 'Có nguy cơ gây nguy hiểm cho người thao tác (thao tác máy móc hoặc lắp ráp) mà không có dấu hiệu báo trước' },
  { scope: 'process', rank: 9, category: 'Không thể thỏa mãn hạng mục yêu cầu của an toàn và/hoặc quy định',
    text: 'Có nguy cơ gây nguy hiểm cho người thao tác (thao tác máy móc hoặc lắp ráp) có dấu hiệu báo trước' },
  { scope: 'process', rank: 8, category: 'Cản trở nghiêm trọng',
    text: 'Phải hủy tất cả sản phẩm. Dừng thao tác của dây chuyền hoặc dừng xuất hàng.' },
  { scope: 'process', rank: 7, category: 'Cản trở lớn',
    text: 'Phải hủy một số sản phẩm sản xuất. Tách khỏi công đoạn chính. Trong đó, bao gồm việc giảm tốc độ của dây chuyền và bổ sung thêm người' },
  { scope: 'process', rank: 6, category: 'Cản trở mức độ trung bình',
    text: 'Tất cả các sản phẩm sản xuất phải sửa ở ngoài dây chuyền, tuy nhiên, có thể chấp nhận được' },
  { scope: 'process', rank: 5, category: 'Cản trở mức độ trung bình',
    text: 'Một số sản phẩm sản xuất phải sửa ở ngoài dây chuyền, tuy nhiên, có thể chấp nhận được' },
  { scope: 'process', rank: 4, category: 'Cản trở mức độ trung bình',
    text: 'Tất cả sản phẩm sản xuất phải sửa tại hiện trường trước khi gia công' },
  { scope: 'process', rank: 3, category: 'Cản trở mức độ trung bình',
    text: 'Một số sản phẩm sản xuất phải sửa tại hiện trường trước khi gia công' },
  { scope: 'process', rank: 2, category: 'Cản trở mức độ nhẹ',
    text: 'Sự bất tiện nhỏ cho công đoạn, thao tác hoặc nhân viên thao tác' },
  { scope: 'process', rank: 1, category: 'Không có ảnh hưởng',
    text: 'Không có ảnh hưởng mà có thể nhận thấy được' },
];
