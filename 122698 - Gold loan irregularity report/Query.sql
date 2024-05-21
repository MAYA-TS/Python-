select m.*, c.emp_name irregularity_entered_empname, d.post_name
  from (select a.zonal_name zone,
               a.region_name region,
               a.inventory inventory_number,
               a.loss loss_amount,
               (count(verified_by) over(partition by inventory))Number_of_verifications_in_that_inventory,
               a.tra_dt irregularity_entered_date,
               a.branch_name,
               a.cust_id customer_id,
               b.tra_dt inventory_creation_date,
               a.pledge_no,
               a.pledge_val pledge_amount,
               b.sticker_no,
               a.gross_weight weight,
               a.Irregularity,
               b.verified_by,
               b.verification_dt,
               a.reportedby irregularity_entered_empcode
               
          from TABLEAU_AUDIT_IRR_RPT a
          left outer join (select a.pledge_no,
                                 a.sticker_no,
                                 trunc(a.tra_dt) verification_dt,
                                 a.usr verified_by,
                                 b.inv_id,
                                 b.tra_dt
                            from mana0809.goldloan_sticker_mst@uatr_backup2 a
                            left outer join (select inv_id, plgno, tra_dt
                                              from mana0809.tbl_gln_inventory_master@uatr_backup2) b
                              on a.pledge_no = b.plgno
                           where post_id in (-252, 821, 309, 248)
                             and a.status = 2
                             and trunc(a.tra_dt) between
                                 trunc(add_months(trunc(sysdate), -12), 'mm') and
                                 trunc(sysdate - 1)) b
            on a.inventory = b.inv_id
         where trunc(a.tra_dt) between add_months(trunc(sysdate), -12) and
               trunc(sysdate - 1)
           and verification_dt < trunc(a.tra_dt)) m
  left outer join (select emp_code, emp_name, post_id
                     from mana0809.emp_master@uatr_backup2) c
    on m.irregularity_entered_empcode = c.emp_code
  left outer join (select a.emp_code, a.emp_name, a.post_id, b.post_name
                     from mana0809.emp_master@uatr_backup2 a,
                          mana0809.post_mst@uatr_backup2   b
                    where a.post_id = b.post_id) d
    on m.verified_by = d.emp_code
