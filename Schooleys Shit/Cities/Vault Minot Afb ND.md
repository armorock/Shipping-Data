# Vault Minot Afb, North Dakota

## All Jobs
```dataview
TABLE job_code, job_name, year, customer
FROM "Job Codes"
WHERE contains(string(city), "Vault Minot Afb") AND contains(string(state), "North Dakota")
SORT year DESC
```

## Structure Count by Product Type
> *Will populate once Shipped Items data is added — counts drawn from the Qty column in each job's Shipped Items table.*
