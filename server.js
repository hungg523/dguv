require('dotenv').config();
const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const { Pool } = require('pg');
const excel = require('exceljs');

const app = express();
app.use(cors());
app.use(bodyParser.json());

// Auth Middleware
const apiAuth = (req, res, next) => {
    if (req.headers['x-auth-token'] === 'DGUV_SECRET_TOKEN_2026') {
        return next();
    }
    return res.status(401).json({ success: false, message: 'Unauthorized' });
};

app.post('/api/login', (req, res) => {
    const { username, password } = req.body;
    if (username === 'MLG2' && password === 'MLG2@2026') {
        return res.json({ success: true, token: 'DGUV_SECRET_TOKEN_2026' });
    }
    res.status(401).json({ success: false, message: 'Sai tài khoản hoặc mật khẩu' });
});

// Phục vụ các file tĩnh (HTML, CSS, JS) trong thư mục hiện tại
app.use(express.static(__dirname));

// Khởi tạo kết nối PostgreSQL bằng chuỗi connection string
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: {
        rejectUnauthorized: false // Bắt buộc cho đa số cloud database (Neon, Supabase)
    }
});

app.post('/api/candidates', apiAuth, async (req, res) => {
    const client = await pool.connect();
    try {
        const { candidateName, interviewerName, interviewDate, t1_score, t2_score, t3_score, t4_score, total_score, final_level, details } = req.body;
        
        await client.query('BEGIN');

        // Insert vào bảng Candidates
        const insertCandidateQuery = `
            INSERT INTO Candidates (CandidateName, InterviewerName, InterviewDate, T1_Score, T2_Score, T3_Score, T4_Score, TotalScore, FinalLevel)
            VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
            RETURNING Id;
        `;
        const candidateValues = [candidateName, interviewerName, interviewDate, t1_score, t2_score, t3_score, t4_score, total_score, final_level];
        const result = await client.query(insertCandidateQuery, candidateValues);
        const candidateId = result.rows[0].id;

        // Insert vào bảng CandidateDetails
        if (details && details.length > 0) {
            const insertDetailQuery = `
                INSERT INTO CandidateDetails (CandidateId, TangName, CompetencyName, Score, Description)
                VALUES ($1, $2, $3, $4, $5);
            `;
            for (const d of details) {
                await client.query(insertDetailQuery, [candidateId, d.TangName, d.CompetencyName, d.Score, d.Description]);
            }
        }

        await client.query('COMMIT');
        res.json({ success: true, message: 'Lưu dữ liệu thành công!', id: candidateId });
    } catch (err) {
        await client.query('ROLLBACK');
        console.error('Lỗi khi lưu DB:', err);
        res.status(500).json({ success: false, message: err.message });
    } finally {
        client.release();
    }
});

app.get('/api/candidates', apiAuth, async (req, res) => {
    try {
        const result = await pool.query('SELECT * FROM Candidates ORDER BY InterviewDate DESC, Id DESC');
        res.json(result.rows);
    } catch (err) {
        console.error('Lỗi khi lấy DB:', err);
        res.status(500).json({ success: false, message: err.message });
    }
});

app.post('/api/export', apiAuth, async (req, res) => {
    try {
        const { candidateIds } = req.body;
        if (!candidateIds || candidateIds.length === 0) {
            return res.status(400).json({ success: false, message: 'Danh sách ID trống' });
        }

        // Tạo chuỗi placeholders $1, $2, ...
        const placeholders = candidateIds.map((_, i) => `$${i + 1}`).join(',');
        const query = `
            SELECT * FROM Candidates 
            WHERE Id IN (${placeholders}) 
            ORDER BY InterviewDate DESC, Id DESC
        `;
        
        const result = await pool.query(query, candidateIds);
        const candidates = result.rows;

        const workbook = new excel.Workbook();
        const sheet = workbook.addWorksheet('Danh Sách Ứng Viên');
        
        // Add Benchmark Table
        sheet.addRow(['BẢNG ĐIỂM CHUẨN SO SÁNH']);
        sheet.mergeCells('A1:E1');
        sheet.getRow(1).font = { bold: true, size: 14, color: { argb: 'FF2563EB' } };
        sheet.addRow([]);
        
        const benchHeaders = sheet.addRow(['Tiêu Chí', 'Staff', 'Supervisor', 'Leader', 'Senior Leader']);
        benchHeaders.font = { bold: true };
        benchHeaders.eachCell(c => c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } });
        
        sheet.addRow(['TẦNG 1: NĂNG LỰC CỐT LÕI', 20, 26, 30, 33]);
        sheet.addRow(['TẦNG 2: NĂNG LỰC CHUYÊN MÔN', 14, 17, 20, 23]);
        sheet.addRow(['TẦNG 3: NĂNG LỰC QUẢN LÝ', 12, 17, 22, 28]);
        sheet.addRow(['TẦNG 4: NĂNG LỰC NHẬN THỨC', 1, 2, 3, 4]);
        const benchTotal = sheet.addRow(['TỔNG ĐIỂM', 47, 62, 75, 88]);
        benchTotal.font = { bold: true };
        benchTotal.getCell(2).font = { bold: true, color: { argb: 'FF2563EB' } };
        benchTotal.getCell(3).font = { bold: true, color: { argb: 'FFDC2626' } };
        benchTotal.getCell(4).font = { bold: true, color: { argb: 'FFD97706' } };
        benchTotal.getCell(5).font = { bold: true, color: { argb: 'FF16A34A' } };

        for(let i=3; i<=8; i++) {
            sheet.getRow(i).eachCell({ includeEmpty: true }, cell => {
                cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            });
        }

        sheet.addRow([]);
        sheet.addRow(['DANH SÁCH ỨNG VIÊN ĐÁNH GIÁ']);
        sheet.mergeCells(`A10:J10`);
        sheet.getRow(10).font = { bold: true, size: 14, color: { argb: 'FF2563EB' } };
        sheet.addRow([]);
        
        // Define columns
        const headerRow = sheet.addRow(['STT', 'Tên Ứng Viên', 'Người PV', 'Ngày PV', 'Tầng 1', 'Tầng 2', 'Tầng 3', 'Tầng 4', 'Tổng Điểm', 'Đánh Giá Phù Hợp']);
        
        sheet.columns = [
            { key: 'stt', width: 8 },
            { key: 'name', width: 25 },
            { key: 'interviewer', width: 20 },
            { key: 'date', width: 15 },
            { key: 't1', width: 12 },
            { key: 't2', width: 12 },
            { key: 't3', width: 12 },
            { key: 't4', width: 12 },
            { key: 'total', width: 15 },
            { key: 'level', width: 20 }
        ];

        // Add rows
        candidates.forEach((row, index) => {
            const dateStr = new Date(row.interviewdate).toLocaleDateString('vi-VN');
            sheet.addRow({
                stt: index + 1,
                name: row.candidatename,
                interviewer: row.interviewername,
                date: dateStr,
                t1: row.t1_score,
                t2: row.t2_score,
                t3: row.t3_score,
                t4: row.t4_score,
                total: row.totalscore,
                level: row.finallevel
            });
        });

        // TÍNH VÀ THÊM DÒNG ĐIỂM TRUNG BÌNH
        if (candidates.length > 0) {
            let sumT1 = 0, sumT2 = 0, sumT3 = 0, sumT4 = 0, sumTotal = 0;
            candidates.forEach(row => {
                sumT1 += row.t1_score || 0;
                sumT2 += row.t2_score || 0;
                sumT3 += row.t3_score || 0;
                sumT4 += row.t4_score || 0;
                sumTotal += row.totalscore || 0;
            });
            
            const n = candidates.length;
            const avgTotal = parseFloat((sumTotal / n).toFixed(1));
            
            // Tính Vị trí phù hợp dựa trên điểm trung bình tổng
            let finalLevel = "Staff";
            let minDiff = Math.abs(avgTotal - 47);
            
            const checkLevel = (target, level) => {
                const diff = Math.abs(avgTotal - target);
                if (diff < minDiff) {
                    minDiff = diff;
                    finalLevel = level;
                }
            };
            
            checkLevel(62, "Supervisor");
            checkLevel(75, "Leader");
            checkLevel(88, "Senior Leader");

            const avgRow = sheet.addRow({
                stt: '',
                name: 'ĐIỂM TRUNG BÌNH CÁC ỨNG VIÊN ĐƯỢC CHỌN',
                interviewer: '',
                date: '',
                t1: parseFloat((sumT1 / n).toFixed(1)),
                t2: parseFloat((sumT2 / n).toFixed(1)),
                t3: parseFloat((sumT3 / n).toFixed(1)),
                t4: parseFloat((sumT4 / n).toFixed(1)),
                total: avgTotal,
                level: finalLevel
            });
            
            avgRow.font = { bold: true, color: { argb: 'FF2563EB' } };
            avgRow.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE2E8F0' }
            };
            
            sheet.mergeCells(`B${avgRow.number}:D${avgRow.number}`);
            avgRow.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
        }

        // Styling candidates table
        headerRow.eachCell((cell) => {
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF2563EB' }
            };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        sheet.eachRow((row, rowNumber) => {
            if (rowNumber > 12) { // Candidate data starts at row 13
                row.eachCell((cell, colNumber) => {
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                    if ([1, 4, 5, 6, 7, 8, 9].includes(colNumber)) {
                        cell.alignment = { horizontal: 'center' };
                    }
                    if (colNumber === 10) {
                        cell.font = { bold: true, color: { argb: 'FF10B981' } };
                        cell.alignment = { horizontal: 'center' };
                    }
                });
            }
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=KetQua.xlsx');

        await workbook.xlsx.write(res);
        res.end();

    } catch (err) {
        console.error('Lỗi khi xuất excel:', err);
        res.status(500).json({ success: false, message: err.message });
    }
});

app.delete('/api/candidates', apiAuth, async (req, res) => {
    try {
        const { candidateIds } = req.body;
        if (!candidateIds || candidateIds.length === 0) {
            return res.status(400).json({ success: false, message: 'Danh sách ID trống' });
        }
        
        // Vì bảng CandidateDetails có ON DELETE CASCADE, nên xóa bên Candidates nó sẽ tự xóa bản ghi con
        const placeholders = candidateIds.map((_, i) => `$${i + 1}`).join(',');
        const query = `DELETE FROM Candidates WHERE Id IN (${placeholders})`;
        await pool.query(query, candidateIds);
        
        res.json({ success: true, message: 'Đã xóa thành công' });
    } catch (err) {
        console.error('Lỗi khi xóa:', err);
        res.status(500).json({ success: false, message: err.message });
    }
});

const PORT = process.env.PORT || 3111;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
