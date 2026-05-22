# Pennsylvania

## All Jobs
```dataview
TABLE job_code, job_name, city, year, customer
FROM "Job Codes"
WHERE contains(string(state), "Pennsylvania")
SORT year DESC
```

## Cities in Pennsylvania
```dataview
TABLE WITHOUT ID city as "City", length(rows) as "Jobs"
FROM "Job Codes"
WHERE contains(string(state), "Pennsylvania") AND city
GROUP BY city
SORT length(rows) DESC
```

## Structure Count by Product Type
> *Will populate once Shipped Items data is added — counts drawn from the Qty column in each job's Shipped Items table.*
