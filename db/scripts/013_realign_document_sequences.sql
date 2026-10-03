USE BTA_OSG_DB;
GO

-- Offline seats mint their own DocCode, and until now replaying one never told the shared
-- counter, so tbl_DocumentSequences could sit below the numbers already on file. The next
-- connected registration would then re-mint a code that exists, fail the unique key, and drop
-- silently to the local cache. This raises each counter to the highest number actually on file
-- for that type and year. Idempotent, and it never lowers a counter.
DECLARE @highest table (
    DocumentTypeCode VARCHAR(50),
    SequenceYear SMALLINT,
    HighestOnFile INT,
    PRIMARY KEY (DocumentTypeCode, SequenceYear)
);

-- DocCode is PREFIX-yyyy-number, so the number starts two positions past the year block.
INSERT INTO @highest (DocumentTypeCode, SequenceYear, HighestOnFile)
SELECT s.DocumentTypeCode,
       s.SequenceYear,
       MAX(TRY_CAST(SUBSTRING(d.DocCode, LEN(s.DocumentTypeCode) + 7, 20) AS int))
FROM dbo.tbl_DocumentSequences s
JOIN dbo.tbl_Documents d
  ON d.DocCode LIKE s.DocumentTypeCode + '-' + CAST(s.SequenceYear AS varchar(4)) + '-%'
GROUP BY s.DocumentTypeCode, s.SequenceYear;

UPDATE s
SET s.LastNumber = h.HighestOnFile
FROM dbo.tbl_DocumentSequences s
JOIN @highest h
  ON h.DocumentTypeCode = s.DocumentTypeCode
 AND h.SequenceYear = s.SequenceYear
WHERE s.LastNumber < h.HighestOnFile;
GO
