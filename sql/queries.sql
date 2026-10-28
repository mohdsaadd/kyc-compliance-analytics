-- Queries are separated by "-- name: <query_name>" markers and are run by src/database.py.
-- Named parameters: :as_of (ISO date) and :sla_days (integer).

-- name: status_summary
SELECT status,
       COUNT(*) AS cases,
       ROUND(100.0 * COUNT(*) / (SELECT COUNT(*) FROM cases), 1) AS pct_of_total
FROM cases
GROUP BY status
ORDER BY cases DESC;

-- name: turnaround_by_document
SELECT document_type,
       COUNT(*) AS closed_cases,
       ROUND(AVG(julianday(completed_date) - julianday(submitted_date)), 1) AS avg_turnaround_days,
       MAX(CAST(julianday(completed_date) - julianday(submitted_date) AS INTEGER)) AS max_turnaround_days
FROM cases
WHERE completed_date IS NOT NULL
GROUP BY document_type
ORDER BY avg_turnaround_days DESC;

-- name: analyst_performance
SELECT analyst,
       COUNT(*) AS total_cases,
       SUM(CASE WHEN status = 'Approved' THEN 1 ELSE 0 END) AS approved,
       SUM(CASE WHEN status = 'Rejected' THEN 1 ELSE 0 END) AS rejected,
       SUM(CASE WHEN status = 'Pending'  THEN 1 ELSE 0 END) AS pending,
       ROUND(AVG(julianday(completed_date) - julianday(submitted_date)), 1) AS avg_turnaround_days
FROM cases
GROUP BY analyst
ORDER BY total_cases DESC;

-- name: monthly_trend
SELECT strftime('%Y-%m', submitted_date) AS month,
       COUNT(*) AS submitted,
       SUM(CASE WHEN status = 'Approved' THEN 1 ELSE 0 END) AS approved,
       SUM(CASE WHEN status = 'Rejected' THEN 1 ELSE 0 END) AS rejected,
       SUM(CASE WHEN status = 'Pending'  THEN 1 ELSE 0 END) AS pending
FROM cases
GROUP BY month
ORDER BY month;

-- name: overdue_cases
SELECT case_id, customer_name, document_type, region, risk_level, analyst,
       submitted_date,
       CAST(julianday(:as_of) - julianday(submitted_date) AS INTEGER) AS days_pending
FROM cases
WHERE status = 'Pending'
  AND julianday(:as_of) - julianday(submitted_date) > :sla_days
ORDER BY
  CASE risk_level WHEN 'High' THEN 1 WHEN 'Medium' THEN 2 ELSE 3 END,
  days_pending DESC;
