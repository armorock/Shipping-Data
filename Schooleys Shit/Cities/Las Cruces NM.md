# Las Cruces, New Mexico

## All Jobs
```dataview
TABLE job_code, job_name, year, customer
FROM "Job Codes"
WHERE contains(string(city), "Las Cruces") AND contains(string(state), "New Mexico")
SORT year DESC
```

## Structure Count by Product Type
> *Will populate once Shipped Items data is added — counts drawn from the Qty column in each job's Shipped Items table.*
