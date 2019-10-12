<?php
##### Clean date
$row['member_dob'] = date('Y-m-d', strtotime($row['member_dob']));
$row['policy_issue_date'] = date('Y-m-d', strtotime($row['policy_issue_date']));
$row['policy_declare_date'] = date('Y-m-d', strtotime($row['policy_declare_date']));
$row['policy_effective_date'] = date('Y-m-d', strtotime($row['policy_effective_date']));
$row['policy_effective_date_card'] = date('Y-m-d', strtotime($row['policy_effective_date_card']));
$row['policy_expiry_date'] = date('Y-m-d', strtotime($row['policy_expiry_date']));
$row['policy_lapse_date'] = date('Y-m-d', strtotime($row['policy_lapse_date']));
$row['policy_revival_date'] = date('Y-m-d', strtotime($row['policy_revival_date']));
$row['policy_termination_date'] = date('Y-m-d', strtotime($row['policy_termination_date']));
$row['plan_attach_date'] = date('Y-m-d', strtotime($row['plan_attach_date']));
$row['plan_expiry_date'] = date('Y-m-d', strtotime($row['plan_expiry_date']));
$row['rider_attach_date'] = date('Y-m-d', strtotime($row['rider_attach_date']));
$row['rider_expiry_date'] = date('Y-m-d', strtotime($row['rider_expiry_date']));
$row['create_date'] = date('Y-m-d', strtotime($row['create_date']));
$row['edit_date'] = date('Y-m-d', strtotime($row['edit_date']));

