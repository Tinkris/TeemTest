USE VideoVSTO
GO

CREATE PROC GetTable10
AS
BEGIN
	DECLARE @code INT
	DECLARE @type VARCHAR(60)
	DECLARE @cat1 INT
	DECLARE @cat2 INT
	DECLARE @cat3 INT
	DECLARE @total INT
	
	DECLARE TCursor CURSOR FOR 
	SELECT *
	FROM dbo.EgpTypes
	
	CREATE TABLE Table_10
	(
		CodeName VARCHAR(60) NOT NULL,
		Category1 INT Not null,
		Category2 INT Not null,
		Category3 INT Not null,
		Total INT
	);
	
	OPEN TCursor
	FETCH NEXT FROM TCursor INTO @code, @type
	WHILE @@FETCH_STATUS = 0
	BEGIN
		--IF @cat = 3 SET @price = @price + 1000
		--INSERT INTO NewProducts VALUES (@name, @cat, @price)
		--FETCH NEXT FROM TCursor INTO @name, @cat, @price
		SELECT @cat1 = COUNT(*)
		FROM dbo.Egp e1
		WHERE e1.Code = @code 
		AND e1.Category = 1
		
		SELECT @cat2 = COUNT(*)
		FROM dbo.Egp e2
		WHERE e2.Code = @code 
		AND e2.Category = 2
		
		SELECT @cat3 = COUNT(*)
		FROM dbo.Egp e3
		WHERE e3.Code = @code 
		AND e3.Category = 3
		
		SET @total = @cat1 + @cat2 + @cat3
		INSERT INTO Table_10 VALUES (@type, @cat1, @cat2, @cat3, @total)
		FETCH NEXT FROM TCursor INTO @code, @type
	END
	CLOSE TCursor
	DEALLOCATE TCursor
	SELECT * FROM Table_10

END
GO
