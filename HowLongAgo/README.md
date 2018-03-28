# Get text for a date that reads like "a minute ago"

This is a UDF that will take a cell that has a date (with or without a time value), determine the amount of time since that date has occurred, and output a friendly amount of time since that date.

## Examples

```vba
' assuming today is 4/1/2018 10:00:00 AM

=HowLongAgo("1/1/2018")
' returns "3 months ago"

=HowLongAgo("4/1/2018 9:59:00 AM")
' returns "1 minute ago"

=HowLongAgo("4/1/2018 11:00:00 AM") ' (future datetime)
' returns "Hasn't happened yet..."
```

# Methodology

The idea here is to start off with the largest unit (years) and work our way down to the smallest unit (seconds). If the difference in years is 0, for example, we move to finding out what the difference in months is and so on.