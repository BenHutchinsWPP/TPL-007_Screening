# Overview
This library is intended as a job-aid to those performing TPL-007 GeoMagnetic Disturbance (GMD) studies and reviewing GMD data quality. However, all code is provided as-is with no warranty. USE AT YOUR OWN RISK. 

# 0. Input Data
Important data for a GMD assessment includes:
- Substation
  - `Latitude`
  - `Longitude`
  - `Rground`
- Bus
  - `SubNumber`
  - Note: If a Substation with Lat/Long is defined, Lat/Long from the Substation is used. 
- Line
  - `GICUSEPFR`, Use Custom R (Yes or No. Yes means to use the custom entered DC Resistance.)
  - `GICCUSTOMR1`, R (DC)
- LineShunt
  - `GICResistance`, R (DC)
- Shunt
  - `GICResistance`, R (DC)
  - Note: FYI in WECC PSLF data, there may also be a field called `rdcgrd` (Shunt substation ground resistance) which may be treated differently than in PowerWorld. 
- Transformer, GICXFormer
  - `XFCoreType`, Options:
    - `Single Phase`
    - `Three Phase Shell Generic`
    - `3-legged, Three Phase`
    - `5-legged, Three Phase`
    - `7-legged, Three Phase`
    - `Core, Three Phase Generic`
    - `Unknown`
  - `XFConfiguration`,`XFConfiguration:1`,`XFConfiguration:2`, Winding Configuration Options: 
    - `Wye`
    - `GWye`
    - `Delta`
    - `Shunt Wye`
    - `Shunt GWye`
    - `Shunt Delta`
    - `Unknown`
  - `XFIsAutoXF`, AutoTransformer Options: 
    - `Yes`
    - `No`
    - `Unknown`
  - `GICManualCoilR`, Use Manually Entered Winding Resistances (Yes or No)
  - `GICXFCoilR1`, `GICXFCoilR1:1`, `GICXFCoilR1:2`, Winding Resistances (Ohms)
  - `GICBlockDevice`, GIC Blocked on Transformer Neutral (Yes or No)
  - `GICModelParam`, GIC Model First Segment Slope
    - Per-unit k-factor
  - `GICModelParam:1`, GIC Model Param Break Point
    - Per unit reactive power (Q) breakpoint in 2-segment piecewise linear k-factor
  - `GICModelParam:2`, GIC Model Param Second Segment Slope
    - Per-unit k-factor
  - `GICModelType`, GIC Model Type, Options:
    - `Default`
    - `Piecewise Linear`

# 1. Data Quality Checks
Data quality checks were produced in collaboration between:
- Ben Hutchins (WPP)
- Chris Gilden (Tristate)
- Zach Zornes (Chelan PUD)

Before running analysis, it is best practice to identify errors in the input data first, then attempt to correct the most aggregious errors before proceeding. However, given that any model will have various fields which contain unknowns, missing values, or outliers, it may not be possible to correct all errors across an entire interconnection footprint during the scope of a given study. As such, please use the data checks as a reference as to what may be corrected going forward for data maintenance purposes. It is also suggested that a utility engineer may perform some sensitivity-checks on the input data quality where suspect data occurs. 

Some example sensitivity tests may include:
- For lines with heavily suspect DC R values, those may be estimated based on the PowerFlow R1 value as a sensitivity test to see how it impacts the result. 
- For substations with missing Rground values, setting those to 0.1 ohms. 
- For Lat/Long values which are suspect, attempting corrections to those to see how the results change. 
- Checking wither opening/closing unused GSUs impacts the results. 

## 1.1. _GMD Case Quality Check (WPP).py

This routine includes several case quality metrics for TPL-007 studies. To use it, modify the `pw_fp` PowerWorld filepath (PWB case) and `rep_fp` Report filepath at the top of the script, then run the script. 

| Tab | What's wrong? | How do you fix it? |
|------|----------------|--------------------|
| **Bus - Undefined Sub** | Any bus where there is equipment which may be grounded MUST have a defined substation with a defined Rground. | In the bus record, fill out the substation field, then define your Rground for your substation record. |
| **Bus vs Sub - LatLong** | The bus Lat/Long location is more than 0.5 miles away from the defined Substation. | Ensure the Bus has the same Lat/Long as its defined Substation. |
| **Bus vs Sub - Area** | The bus area number does not match the substation area number. | Verify that the Buses are placed in the correct Substations. |
| **SubNum not in BusNums** | **Example:** Substation #1 contains buses [12001, 12002, 12003].<br>A substation number MUST be found in the list of buses which it contains. | In this example, you could change the Substation number to any of these numbers: 12001, 12002, or 12003. |
| **Sub Missing Rground** | Substation Rground is missing (null), or equal to zero. | Enter a measured or assumed value for Rground. |
| **XFMR with Length** | From/To buses are in different geographic lat/long locations.<br>**Threshold:** 0.5 miles. | You could either:<br>1) Ensure From/To buses are in the same substation, or<br>2) Ensure the From/To have the same Lat/Long. |
| **XFMR Missing Data** | The case has “Unknown” for:<br>• GICCoreType (three-leg, etc)<br>• XFConfiguration (Delta/Wye etc)<br>• GICAutoXF (Yes/No is autotransformer) | Ensure all transformer data is filled out. |
| **Line Length Suspect** | Estimated length (miles) based on Lat/Long is very different from estimated length (miles) based on R/X.<br>**Threshold:** Absolute length difference > 0.5 miles, and ratio > 1.5 or < 0.5. | Verify your branch R, X, B, and Lat/Long of From/To buses. |
| **Line R Suspect** | 1) DC Resistance (GICCUSTOMR1) should be lower than AC Resistance (GICPFR1).<br>2) DC Resistance (GICCUSTOMR1) should be within ~20% of AC Resistance (GICPFR1). | Double-check your entry of R1 (positive sequence R) and Rdc. |
| **Line Changes NomkV** | A line where From/To buses have different nominal kV values.<br>**This is a CRITICAL error as you cannot read the GIC model into GICHarm if you do this!** | Open the branch, or change NominalkV to match. |

### Note: Line R Suspect
`line_r_suspect(acdc_high=1.2, acdc_low=0.99)`

Report Lines with high mismatch between PowerFlow R1 and custom entered DC resistance.
- DC Resistance should be lower than Positive Sequence AC resistance (R1). 
- DC Resistance should be close to Positive Sequence AC resistance (R1) (default threshod: within ~20%)

AC resistance should usually be higher than DC resistance for any given conductor. Additionally, a ratio of AC to DC resistance of more than 1.16 appears to be a significant outlier when looking at the following reference: 

> "Electric Power Distribution Engineering" (Turan Gonen) Second Edition. Appendix A.

## 1.2. Geographic Map Checks
Plotting the buses & lines by their geography on a PowerWorld geographic diagram is suggested. "_Auto Insert 1-line.aux" can be used to create a one-line map of the study area.

Types of checks which are aided by Geographic Maps of Buses & Lines:
- Identify buses which are far away from your interconnection and should not be. 
- Identifying differences in the current base-case, compared to a previous study case. Overlaying both maps on top of eachother, then toggling between them, makes it incredibly obvious where data may have been entered differently (and possibly incorrectly). 

# 2. PowerWorld Analysis

## Area Ignore GIC Losses
In the Area table, there is a variable called `GICIgnoreMvarLosses`. You may find it here:
- Addons &rarr; GIC &rarr; Areas &rarr; GICIgnoreMvarLosses (Ignore GIC Losses)

If this flag is `Yes`, please note that all transformer DC GIC flow values will be shown as zero (0) in the result tables for those areas as well. 

`GIC_Options.aux` sets all areas to `GICIgnoreMvarLosses`=`NO` to ensure no transformers are missed in the study reports. 

## Bus vs Sub Lat/Long
Suppose a bus has a defined Lat/Long which does not match the designated Substation Lat/Long. In this case, PowerWorld takes the Substation record as the main record, and ignore the Bus Lat/Long. The Bus Lat/Long is only used when a Substation with Lat/Long is not defined for that bus. 

To make this behavior explicitly clear, `Autofill_Bus_LatLong_From_Sub.aux` can be used to set all Bus Lat/Long values from the the Substation Lat/Longs, wherever there is a value (and it is not null). 

## GSUs on Open Generators
As per previous study practices, `Open_Unused_GSUs.aux` can open Generation Step-Up (GSU) transformers which serve open generators. 

# Citations
This library uses the TAMU `ACTIVSg2000.PWB` model as sample input to the scripts in this library. 

For more information, please see the [TAMU ACTIVSg2000 Webpage](https://electricgrids.engr.tamu.edu/electric-grid-test-cases/activsg2000/) and the below reference:

[1] A. B. Birchfield; T. Xu; K. M. Gegner; K. S. Shetye; T. J. Overbye, “Grid Structural Characteristics as Validation Criteria for Synthetic Networks,”  in IEEE Transactions on Power Systems, vol. 32, no. 4, pp. 3258-3265, July 2017.

Credit to PowerWorld staff for providing support in developing the WPP methodology for TPL-007, including several scripts used in this repository such as:
- `Open_Unused_GSUs.aux`
- [PW Earth Resistivity Model for GIC Calculations
](https://www.powerworld.com/knowledgea-base/earth-resistivity-model-for-gic-calculations)
  - `NERC_USGS_2017_Regions20220208.aux`
  - `NERC_USGS_2017_RegionsAB10_20220208`
- [PW Transformer Time-Series for NERC Benchmark GMD Event and Supplemental GMD Event
](https://www.powerworld.com/knowledge-base/transformer-time-series-for-nerc-benchmark-gmd-event)
  - `NERC_GMDBenchmarkEventTimeSeries.csv`
  - `NERC_GMDSupplementalEventTimeSeries.csv`
