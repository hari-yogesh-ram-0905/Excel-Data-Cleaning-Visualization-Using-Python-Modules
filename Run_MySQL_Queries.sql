
CREATE DATABASE IF NOT EXISTS Python_Project;

USE Python_Project;

SELECT * FROM cleaned_data
ORDER BY Name;

DROP TABLE IF EXISTS cleaned_data;
