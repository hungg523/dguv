-- Chạy lệnh này trên công cụ quản lý PostgreSQL (ví dụ: Supabase SQL Editor, Neon SQL Editor)

CREATE TABLE IF NOT EXISTS Candidates (
    Id SERIAL PRIMARY KEY,
    CandidateName VARCHAR(255) NOT NULL,
    InterviewerName VARCHAR(255) NOT NULL,
    InterviewDate DATE NOT NULL,
    T1_Score INT NOT NULL,
    T2_Score INT NOT NULL,
    T3_Score INT NOT NULL,
    T4_Score INT NOT NULL,
    TotalScore INT NOT NULL,
    FinalLevel VARCHAR(50) NOT NULL
);

CREATE TABLE IF NOT EXISTS CandidateDetails (
    Id SERIAL PRIMARY KEY,
    CandidateId INT NOT NULL REFERENCES Candidates(Id) ON DELETE CASCADE,
    TangName VARCHAR(255) NOT NULL,
    CompetencyName VARCHAR(255) NOT NULL,
    Score INT NOT NULL,
    Description TEXT
);
