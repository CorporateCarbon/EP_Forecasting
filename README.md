# EP_Abatement

### Important Points ###

# Most FullCAM outputs are from end of month (EOM) -> to end of month
# Therefore when reading it, the system must read from EOM to EOM
# However, reporting periods dont overlap, so there is often a disconect in the way we calculate vs record RP data
# For example; a RP from 1/1/2025 to 31/12/2025 must be calculated as 31/12/2024 to 31/12/2025 and will therefore probably overlap with next door RPs.

### WHY THIS MATTERS ###

    - The system will AUTOMATICALLY index the entered date in EP_Forecast_Runner by -1 day, to catch the previous EOM
    - E.g. 1/1/2025 in the system is turned into 31/12/2024