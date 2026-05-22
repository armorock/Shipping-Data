# West Valley City, Utah

## All Jobs
```dataview
TABLE job_code, job_name, year, customer
FROM "Job Codes"
WHERE contains(string(city), "West Valley City") AND contains(string(state), "Utah")
SORT year DESC
```

## Structure Count by Product Type
> *Will populate once Shipped Items data is added — counts drawn from the Qty column in each job's Shipped Items table.*
