# cigos_code
**Đầu bài: Cigos có danh sách các bài báo (papers) và danh sách các chuyên gia để chấm điểm các bài báo đó (reviewers). Code dùng để thực hiện chia số papers đều cho tất cả các reviewer nhất với các điều kiện như sau :**
1. Mỗi reviewer ko chấm quá số paper mà họ đã yêu cầu (xem cột max of papers to review)
2. Mỗi bài cần ít nhất 2 người chấm, 
3. Mỗi bài từ VN thì cần ít nhất một người ở nước ngoài chấm
4. Xong điều kiện 1, 2 và 3 mới phân số người có thể chấm thêm để mỗi bài có được 3 người chấm, đến khi nào hết khả năng thì thôi

***
## Hướng dẫn chạy code:
**Phân bổ reviewer chấm paper:**
1. Kiểm tra và thay đổi 2 file đầu vào với đúng các cột như file mẫu
- File **reviewers.xls** trong thư mục file
- File **paper.xls** trong thư mục file
2. Kết quả là file **Format_editors.xlsx** trong thư mục file/final


***
## Giải thích các file:
**PhanBoPaperLogic1Coppy.java:**
- File này là file phân bổ với logic được update mới nhất
- Chạy file PhanBoPaperLogic1Coppy.java bằng cách run Java File
- Có rất nhiều cách để chạy được Java File (Có thể GG để tìm cách hợp với bạn nhất)
- Như cách mình dùng là mình dùng IDE Intellij nên chỉ cần chuột phải và chọn run file thôi là nó chạy 

**LogicDoUuTienTopic.java:**
- Sắp xếp độ ưu tiên của các TOPIC

**Các file khác không cần quan tâm ạ, là mấy file em test logic thôi**
