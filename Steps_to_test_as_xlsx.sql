

--chatgpt notes
https://chatgpt.com/c/688713da-9118-800a-9f1d-97d6ae8f59bc
/*step:1 */
--connect as sysdba and create and grant below permissions..
GRANT EXECUTE ON UTL_FILE TO yourschemaname;
CREATE OR REPLACE DIRECTORY FILE_DIR AS 'E:\Filedir';
GRANT READ, WRITE ON DIRECTORY FILE_DIR TO XXCCT;
SELECT * FROM all_directories WHERE directory_name = 'FILE_DIR';
/
--test file is creating in your dir or not..
BEGIN
  -- Write Excel to E:\Filedir\report.xlsx
  as_xlsx.save(p_directory => 'FILE_DIR', p_filename => 'report.xlsx');
END;
/

DECLARE
  f UTL_FILE.FILE_TYPE;
BEGIN
  f := UTL_FILE.FOPEN('FILE_DIR', 'test.txt', 'W');
  UTL_FILE.PUT_LINE(f, 'Test write success');
  UTL_FILE.FCLOSE(f);
END;
/

--then exec below to create your excel sheet in sheet 2.
BEGIN
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet('EMP Export');
  as_xlsx.query2sheet(
    p_sql            => 'SELECT empno, ename, job, sal, deptno FROM emp',
    p_column_headers => TRUE,
    p_directory      => 'FILE_DIR',
    p_filename       => 'emp_export_1.xlsx',
    p_autofilter     => TRUE,
    p_table_style    => 'TableStyleMedium2'
  );
END;
/
  --WORKING AS EXPECTED
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet(p_sheetname=>'emp2',p_show_gridlines=>TRUE,p_show_headers=>TRUE);
  as_xlsx.query2sheet(
    p_sql            => 'SELECT empno, ename, job, sal, deptno FROM emp',
    p_column_headers => TRUE,
    p_autofilter     => TRUE,
    p_sheet=> 2,
    p_table_style    => 'TableStyleMedium2'
  );
  as_xlsx.save( 'FILE_DIR', 'EMP REPORT.xlsx' );
end;
 /

--step2
--adding smtp server 
--https://github.com/ChangemakerStudios/Papercut-SMTP/releases
