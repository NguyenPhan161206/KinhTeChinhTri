# BÁO CÁO MÔ TẢ CÔNG THỨC VÀ ĐIỀU KIỆN MÔ PHỎNG TÀI CHÍNH
**Dự án:** AI Flow Corp (SaaS Model)
**Tệp nguồn:** `financial_model_Final.xlsx`

## I. CÁC THAM SỐ GIẢ ĐỊNH (ASSUMPTIONS)
- Vốn đầu tư ban đầu (K): 2.000.000.000 VNĐ
- Tỷ lệ tăng trưởng/tháng: 5% (0.05)
- Lương bình quân: [lương bình quân] VNĐ/người/tháng
- Nhân sự ban đầu: 30 người
- Tuyển mới hàng tháng: 1 người

## II. HỆ THỐNG KÝ HIỆU VÀ CÔNG THỨC
Dựa trên lý luận Kinh tế Chính trị Mác-Lênin:

1. **Số nhân sự (NS):**
   `NS(n) = NS(n-1) + 1`
   *(Mở rộng quy mô tư bản khả biến v)*

2. **Doanh thu / Giá trị doanh nghiệp (DT):**
   `DT(n) = DT(n-1) * 1.05`
   *(Tăng trưởng nhờ tối ưu năng suất lao động bằng AI)*

3. **Lãi tháng (lãi):**
   `lãi(n) = DT(n-1) * 0.05`
   *(Giá trị thặng dư m tạo ra trong kỳ)*

4. **Tổng lương chi (TLC):**
   `TLC(n) = NS(n) * [lương bình quân]`
   *(Chi phí tái sản xuất sức lao động v)*

5. **Số dư vốn (SDV):**
   `SDV(n) = SDV(n-1) + lãi(n) - TLC(n)`
   *(Khả năng tích lũy tư bản và duy trì dòng tiền)*

## III. ĐIỀU KIỆN MÔ PHỎNG
1. **Giá trị thặng dư tương đối:** Tăng trưởng 5% dựa trên việc áp dụng công nghệ AI để giảm thời gian lao động tất yếu.
2. **Tích lũy tư bản:** 50% thặng dư được tái đầu tư vào hạ tầng (c) và 50% vào nhân sự (v).
3. **Rủi ro lồng ghép:** Bao gồm các biến động về nhân sự (thai sản), thị trường (mất đơn hàng) và cạnh tranh (săn đầu người) tại các mốc tháng 3, 7 và 10.
