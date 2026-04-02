function calculateFinancials(month) {
    const d = window.staffData[month];
    
    // 1. Tổng giá trị tạo ra (G = c + v + m)
    // Tính toán m dựa trên năng suất lao động (v)
    let totalValue = (d.sales * 60) + (d.mkt * 60) + (d.tech * 120) + (d.ops * 100);
    
    // 2. Chi phí tư bản khả biến (v - Tiền lương)
    let totalSalary = (d.sales * 8) + (d.mkt * 12) + (d.tech * 23) + (d.ops * 11) + (d.acc * 12);
    
    // Tác động vĩ mô (Lạm phát)
    if(window.macroState.inflation === 1) totalSalary *= 1.15;
    
    // Rủi ro hệ thống (Tháng 3, 7, 10)
    if (month === 3) totalValue *= 0.85;
    if (month === 7) totalValue *= 0.85;
    if (month === 10) totalSalary *= 1.1;

    // 3. Tổng giá trị thặng dư tạo ra (m)
    const totalSurplus = totalValue - totalSalary;

    // 4. ĐÒN BẨY TÀI CHÍNH (Sử dụng vốn vay 4 Tỷ)
    const loanAmount = 4000; // 4.000 Tr VNĐ (4 Tỷ)
    const monthlyInterestRate = 0.01; // 1% / tháng
    
    // Lợi tức (z) - Phần m trả cho Ngân hàng
    const interestZ = loanAmount * monthlyInterestRate;
    
    // 5. Lợi nhuận doanh nghiệp (m_dn = m - z)
    // Nếu kịch bản lãi suất cao, z tăng lên gấp đôi
    const actualInterestZ = window.macroState.interest === 1 ? interestZ * 2 : interestZ;
    const enterpriseProfit = totalSurplus - actualInterestZ;

    return { 
        totalStaff: d.sales + d.mkt + d.tech + d.ops + d.acc, 
        mVal: enterpriseProfit, // Trả về lợi nhuận sau khi trừ lãi vay
        totalSurplus: totalSurplus, // Tổng thặng dư ban đầu
        interestZ: actualInterestZ, // Phần m trả cho chủ nợ
        totalSalary, 
        cVal: 1400 + (month * 40) 
    };
}