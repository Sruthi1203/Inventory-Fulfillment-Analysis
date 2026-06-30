# Inventory & Sales Order Fulfillment Analysis

## Project Overview

This project simulates a real-world inventory operations workflow to determine whether open sales orders can be fulfilled based on available inventory, item status, kit components, and incoming purchase orders.

The objective was to build a decision-support system that automatically recommends the appropriate operational action for each sales order while providing clear explanations for business users.

---

## Objectives

- Analyze open sales orders against available inventory
- Identify inventory shortages at the requested warehouse
- Recommend operational actions for each order
- Evaluate inventory availability across multiple warehouse locations
- Validate kit/package items using component-level inventory
- Avoid duplicate procurement by checking existing purchase orders
- Create user-friendly reporting and lookup tools

---

## Dataset

The project integrates multiple operational datasets:

- Open Sales Orders
- On-Hand Inventory
- Kit Components
- Open Purchase Orders

---

## Business Rules Implemented

### Ready to Fulfill
Order can be shipped immediately using inventory available at the requested location.

### A – Place Purchase Order
Inventory is unavailable and the item is eligible for replenishment.

### B – Transfer Inventory
Inventory is unavailable at the requested location but exists at another warehouse.

### C – Cancel Order
Applied to items marked as:
- Discontinued
- Sell Down and Discontinue

### Existing PO – Monitor
An open purchase order already exists for the item, preventing duplicate procurement.

---

## Workflow

### 1. Inventory Validation

- Compared Sales Order quantity with available inventory at the requested location.
- Calculated shortages.

---

### 2. Multi-Location Inventory Check

If inventory was unavailable locally:

- Calculated total inventory across all warehouse locations.
- Determined whether inventory transfer was possible.

---

### 3. Item Status Validation

Integrated item master data to determine:

- Active
- Custom
- Development
- Mass
- Discontinued
- Sell Down and Discontinue

Business rules were applied accordingly.

---

### 4. Open Purchase Order Validation

Checked whether incoming inventory already existed.

This prevented recommending duplicate purchase orders.

---

### 5. Kit Component Validation

Kit items were evaluated separately.

Since kits consist of multiple components:

- Each component inventory was validated.
- Required quantities were calculated.
- Minimum kits possible were determined based on the limiting component.

---

## Features

### Sales Order Fulfillment Analysis

Main worksheet performing:

- Inventory validation
- Shortage calculation
- Item status lookup
- Final action recommendation
- Notes generation
- Purchase order validation

---

### Order Fulfillment Dashboard

Interactive summary including:

- Action distribution
- Fulfillment statistics
- Percentage breakdown
- Pivot table analysis

---

### Sales Order Lookup Tool

Built an interactive lookup utility allowing users to:

- Search by Document Number
- View inventory availability
- View item status
- View recommended action
- View explanatory notes

This simulates real-time operational support.

---

### Kit Validation Sample

Demonstrates:

- Component-level analysis
- Required quantities
- Available inventory
- Kit feasibility calculations

---

## Excel Functions Used

- XLOOKUP
- SUMIF
- SUMIFS
- MINIFS
- Nested IF Logic
- COUNTIF
- Pivot Tables
- Conditional Formatting
- Data Validation
- Dynamic Lookup Formulas

---

## VBA

Created a VBA macro to:

- Clear lookup inputs
- Reset displayed results
- Improve user experience

---

## Challenges

### Kit Components

A single kit maps to multiple components (one-to-many relationship).

Rather than relying on kit inventory alone, fulfillment was determined using component availability.

---

### Open Purchase Orders

Incoming inventory should not trigger another purchase order.

The solution introduced an additional recommendation:

**Existing PO – Monitor**

to prevent duplicate procurement.

---

### Data Quality

Observed:

- Duplicate records
- Missing values
- Inconsistent purchase order data

These observations were documented while preserving original source data.

---

## Future Enhancements

- SKU-based lookup tool
- Full kit lookup support
- Interactive dashboard
- Automated reorder alerts
- Timestamp-based duplicate validation
- Improved naming conventions aligned with business standards
- Power BI dashboard integration
- SQL-backed inventory validation
- Automated data refresh

---

## Skills Demonstrated

- Business Process Analysis
- Inventory Management
- Supply Chain Analytics
- Excel Automation
- Decision Support Systems
- Data Cleaning
- Data Validation
- Dashboard Development
- Problem Solving
- Business Rule Implementation

---

## Tools Used

- Microsoft Excel
- XLOOKUP
- SUMIF / SUMIFS
- MINIFS
- Nested IF Logic
- Pivot Tables
- VBA Macros
- Dynamic Lookup Tool


---

## Key Outcome

Developed a complete inventory fulfillment decision-support solution capable of:

- Identifying fulfillment risks
- Recommending operational actions
- Validating kit feasibility
- Preventing duplicate procurement
- Supporting real-time order lookup
- Presenting executive-level summaries through dashboards
