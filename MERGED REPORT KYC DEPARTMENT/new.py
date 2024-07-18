import xlsxwriter
import os
import pandas as pd
import cx_Oracle
import openpyxl
from pandas import ExcelWriter
from openpyxl import Workbook
from sqlalchemy import create_engine

from email.message import EmailMessage
from email.utils import make_msgid
import mimetypes
import smtplib


conn = cx_Oracle.connect("kpmg", "Asd$1234", "HISTDB")
print("Oracle database connected")

df5=pd.read_sql("""select w.fzm,r.reg_name,r.area_name, r.branch_name, z.*
 from (SELECT req_branch,
  count(request_id) requested_count,
  count(case when status_id=1 then request_id end) approved_count,
  count(case when status_id=2 then request_id end) rejected_count,
  count(case when status_id=0 then request_id end) pending_count,
  COUNT(CASE WHEN  status_id = 1 and time_interval = '1-5' THEN request_id END) AS approved_one_to_five,
  COUNT(CASE WHEN  status_id = 1 and time_interval = '5-10' THEN request_id END) AS approved_five_to_ten,
  COUNT(CASE WHEN  status_id = 1 and time_interval = '10-15' THEN request_id END) AS approved_ten_to_fifteen,
  COUNT(CASE WHEN  status_id = 1 and time_interval = 'above 15 min' THEN request_id END) AS approved_above_fifteen,
  COUNT(CASE WHEN  status_id = 2 and time_interval = '1-5' THEN request_id END) AS rejected_one_to_five,
  COUNT(CASE WHEN  status_id = 2 and time_interval = '5-10' THEN request_id END) AS rejected_five_to_ten,
  COUNT(CASE WHEN  status_id = 2 and time_interval = '10-15' THEN request_id END) AS rejected_ten_to_fifteen,
  COUNT(CASE WHEN  status_id = 2 and time_interval = 'above 15 min' THEN request_id END) AS rejected_above_fifteen
FROM
 ( SELECT
    x.*,
    CASE
      WHEN status_id = 1 AND approved_diff >= 0.1 AND approved_diff <= 5 THEN '1-5'
      WHEN status_id = 1 AND approved_diff > 5 AND approved_diff <= 10 THEN '5-10'
      WHEN status_id = 1 AND approved_diff > 10 AND approved_diff <= 15 THEN '10-15'
      WHEN status_id = 1 AND approved_diff > 15 THEN 'above 15 min'
      WHEN status_id = 2 AND rejected_diff >= 0.1 AND rejected_diff <= 5 THEN '1-5'
      WHEN status_id = 2 AND rejected_diff > 5 AND rejected_diff <= 10 THEN '5-10'
      WHEN status_id = 2 AND rejected_diff > 10 AND rejected_diff <= 15 THEN '10-15'
      WHEN status_id = 2 AND rejected_diff > 15 THEN 'above 15 min'
    END AS time_interval
  FROM (
    SELECT   a.*, 
      case when status_id=1 then 'Approved'
       when status_id=2  then 'Rejected' end as  status,
      ROUND((approved_date - requested_date) * 24 * 60, 1) AS approved_diff,
      ROUND((rejected_date - requested_date) * 24 * 60, 1) AS rejected_diff
    FROM
      mana0809.customer_merge@uatr_backup2 a
    WHERE
      TRUNC(a.requested_date) >= TRUNC(SYSDATE)
         ) x  )-- y  WHERE time_interval IS NOT NULL and status_id in (1,2)
    GROUP BY   req_branch
    ORDER BY  req_branch )z
    left outer join   mana0809.branch_dtl_new@uatr_backup2  r
    on z.req_branch= r.branch_id
    left outer join   mana0809.tbl_fzm_master@uatr_backup2  w
    on r.reg_id= w.region_id """,con=conn)
df6=pd.read_sql("""select x.*,
       case
         when  status_id =1and approved_diff >= 0.1 and approved_diff <= 5 then
          '1-5'
         when  status_id =1and approved_diff > 5 and approved_diff <= 10 then
          '5-10'
         when  status_id =1and approved_diff > 10 and approved_diff <= 15 then
          '10-15'
         when  status_id =1and approved_diff > 15 then
          'above 15 min'
      
         when status_id=2 and rejected_diff >= 0.1 and rejected_diff <= 5 then
          '1-5'
         when status_id=2 and rejected_diff > 5 and rejected_diff <= 10 then
          '5-10'
         when status_id=2 and rejected_diff > 10 and rejected_diff <= 15 then
          '10-15'
         when status_id=2 and rejected_diff > 15 then
          'above 15 min'
       end as time_interval
  from (select a.*,
               round((approved_date - requested_date) * 24 * 60, 1) AS approved_diff,
               round((rejected_date - requested_date) * 24 * 60, 1) AS rejected_diff
          from mana0809.customer_merge@uatr_backup2 a
         where trunc(a.requested_date) >= trunc(sysdate)
         ) x 
""",con=conn)



print("k")
writer = pd.ExcelWriter("MERGING VERIFICATION NEW DASHBOARD.xlsx", engine="openpyxl")
# writer = pd.ExcelWriter("C:/mail_crf/mail/operation_report.xlsx", engine="openpyxl")
# writer = pd.ExcelWriter("C:/Users/357274/Downloads/repot/Phase1.xlsx", engine="openpyxl")
#writer = pd.ExcelWriter("operation_report111.CSV")
df5.to_excel(writer, sheet_name="CONSOLIDATED", index=False)
df6.to_excel(writer, sheet_name="DETAILED", index=False)

writer.save()
