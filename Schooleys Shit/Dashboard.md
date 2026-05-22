# Armorock Shipping Dashboard

## Jobs by State
```dataview
TABLE WITHOUT ID state as "State", length(rows) as "Jobs"
FROM "Job Codes"
WHERE state
GROUP BY state
SORT length(rows) DESC
```

## Jobs by Year
```dataview
TABLE WITHOUT ID year as "Year", length(rows) as "Jobs"
FROM "Job Codes"
WHERE year
GROUP BY year
SORT year DESC
```

## Jobs by Plant
```dataview
TABLE WITHOUT ID plant as "Plant", length(rows) as "Jobs"
FROM "Job Codes"
WHERE plant
GROUP BY plant
SORT length(rows) DESC
```

## Structure Counts by Product Type
> *Will populate once Shipped Items data is added — counts drawn from the Qty column in each job's Shipped Items table.*
