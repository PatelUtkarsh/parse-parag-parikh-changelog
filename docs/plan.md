# Plan to Update Mutual Fund Parsing Script

The goal is to analyze the monthly changes in a mutual fund's portfolio by looking at both the change in the number of shares and the change in the total market value of each holding. This provides a clearer picture of the fund manager's activity and market impact. The report will also include the implied end-of-month share prices to give context on price movements.

- [ ] **Modify `src/ParseCommand.php`**
    - [ ] **Update `get_excel_column_map()` method:**
        - [ ] Ensure logic correctly extracts 'Quantity' and 'Market/Fair Value' columns.
    - [ ] **Update `flattenSectionData()` method:**
        - [ ] Ensure logic correctly aggregates both `quantity` and `market_value`.
    - [ ] **Update `execute()` and `displaySectionWiseComparison()` methods:**
        - [ ] Calculate the difference for both share quantity and market value.
        - [ ] Calculate the implied `Old Price` and `New Price` from the market value and quantity for each period, handling cases where quantity is zero.
        - [ ] Sort the results based on the absolute change in market value to highlight the most significant allocation shifts.
        - [ ] Update the output tables to display `Change (Value)`, `Old Price`, `New Price`, `Change (Shares)`, and `Current %`.
- [ ] **Verify Changes**
    - [ ] Run the script with sample data.
    - [ ] Check that the output tables are correctly sorted and display all the new price and value information.
