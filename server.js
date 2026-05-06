require('dotenv').config();
const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const sql = require('mssql');
const excel = require('exceljs');

const app = express();
app.use(cors());
app.use(bodyParser.json());

// Phục vụ các file tĩnh (HTML, CSS, JS) trong thư mục hiện tại
app.use(express.static(__dirname));

const dbConfig = {
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    server: process.env.DB_SERVER,
    database: process.env.DB_NAME,
    options: {
        encrypt: false,
        trustServerCertificate: true 
    }
};

app.post('/api/candidates', async (req, res) => {
    try {
        const { candidateName, interviewerName, interviewDate, t1_score, t2_score, t3_score, t4_score, total_score, final_level, details } = req.body;
        
        await sql.connect(dbConfig);
        
        // Check if table exists, if not, wait. They should run init.sql
        const result = await sql.query`
            INSERT INTO Candidates (CandidateName, InterviewerName, InterviewDate, T1_Score, T2_Score, T3_Score, T4_Score, TotalScore, FinalLevel)
            OUTPUT INSERTED.Id
            VALUES (${candidateName}, ${interviewerName}, ${interviewDate}, ${t1_score}, ${t2_score}, ${t3_score}, ${t4_score}, ${total_score}, ${final_level})
        `;
        
        const candidateId = result.recordset[0].Id;

        // Insert details
        if (details && details.length > 0) {
            const ps = new sql.PreparedStatement();
            ps.input('candidateId', sql.Int);
            ps.input('competencyName', sql.NVarChar(255));
            ps.input('score', sql.Int);
            ps.input('description', sql.NVarChar(sql.MAX));
            
            await ps.prepare(`INSERT INTO CandidateDetails (CandidateId, CompetencyName, Score, Description) VALUES (@candidateId, @competencyName, @score, @description)`);
            
            for (let detail of details) {
                await ps.execute({
                    candidateId: candidateId,
                    competencyName: detail.competencyName,
                    score: detail.score,
                    description: detail.description
                });
            }
            await ps.unprepare();
        }
        
        res.status(200).json({ success: true, id: candidateId });
    } catch (err) {
        console.error(err);
        res.status(500).json({ success: false, message: err.message });
    } finally {
        // sql.close(); // Not closing so connection pool can be reused, or could close.
    }
});

app.get('/api/candidates', async (req, res) => {
    try {
        await sql.connect(dbConfig);
        const result = await sql.query`SELECT * FROM Candidates ORDER BY InterviewDate DESC, Id DESC`;
        res.json(result.recordset);
    } catch (err) {
        console.error(err);
        res.status(500).json({ success: false, message: err.message });
    }
});

// Endpoint to export selected candidates
app.post('/api/export', async (req, res) => {
    try {
        const { candidateIds } = req.body;
        if (!candidateIds || candidateIds.length === 0) {
            return res.status(400).json({ success: false, message: 'No candidates selected' });
        }

        await sql.connect(dbConfig);
        const idsString = candidateIds.join(',');
        
        const candidates = await sql.query(`SELECT * FROM Candidates WHERE Id IN (${idsString})`);
        
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
        candidates.recordset.forEach((row, index) => {
            const dateStr = new Date(row.InterviewDate).toLocaleDateString('vi-VN');
            sheet.addRow({
                stt: index + 1,
                name: row.CandidateName,
                interviewer: row.InterviewerName,
                date: dateStr,
                t1: row.T1_Score,
                t2: row.T2_Score,
                t3: row.T3_Score,
                t4: row.T4_Score,
                total: row.TotalScore,
                level: row.FinalLevel
            });
        });

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
        res.setHeader('Content-Disposition', 'attachment; filename=' + 'Danh_Sach_Ung_Vien.xlsx');
        
        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        console.error(err);
        res.status(500).json({ success: false, message: err.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
