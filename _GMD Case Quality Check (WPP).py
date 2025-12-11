from pathlib import Path
import pandas as pd
import win32com.client
import PW_Scripts.wpp_lib as wpp_lib
from math import radians, cos, sin, asin, sqrt

cur_dir = Path(__file__).parent

# Update for your own specific case.
pw_fp = cur_dir / 'ACTIVSg2000.PWB'
rep_fp = cur_dir / 'ACTIVSg2000.PWB GMD Case Quality Check.xlsx'
wpp_lib.mva_mismatch_threshold = 1.0

def haversine(lon1, lat1, lon2, lat2):
    """
    Calculate the great circle distance in kilometers between two points 
    on the earth (specified in decimal degrees)
    """
    # convert decimal degrees to radians 
    lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])

    # haversine formula 
    dlon = lon2 - lon1 
    dlat = lat2 - lat1 
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * asin(sqrt(a)) 
    r = 6371 # Radius of earth in kilometers. Use 3956 for miles. Determines return value units.
    return c * r

def run_gmd_quality_checks(SimAuto, pw_fp: Path, rep_fp: Path) -> dict[str,pd.DataFrame]:
    def bus_unmapped_sub() -> pd.DataFrame:
        """
        Reports any non-star buses which do not have a defined substation.
        To do: Filter on buses which have equipment that may have grounds. It only impacts calculations if there's a possibly grounded piece of equipment there (transformer, or shunt)
        """
        table = 'Bus'
        parameter_type: dict[str,type] = {
            'Number': int
            ,'Name': str
            ,'NomkV': float
            ,'SubNumber': int
            ,'AreaNum': int
            ,'OwnerNumber': int
            ,'OwnerName': str
            ,'DataMaintainer': str
            ,'IsStarBus': str
            ,'IsLikelyStarBus': str
        }
        df = wpp_lib.get_param_df(SimAuto, table, parameter_type)
        df = df[
            df['SubNumber'].isnull() & 
            df['IsStarBus'].str.contains('NO') & 
            df['IsLikelyStarBus'].str.contains('NO')
            ]
        df['>200kV'] = df['NomkV'].apply(lambda x: 'YES' if x > 200 else 'NO')
        df = df.sort_values(by=['>200kV','AreaNum','NomkV','Number'], ascending=[False, True, False, True])
        df = df.drop(columns=['SubNumber','IsStarBus','IsLikelyStarBus'])
        return df

    def bus_sub_latlong_mismatch(threshold_miles: float = 0.5) -> pd.DataFrame:
        """Reports buses with a lat/long that doesn't match their sub lat/long within the threshold."""
        table = 'Bus'
        parameter_type: dict[str,type] = {
            'Number': int
            ,'Name': str
            ,'NomkV': float
            ,'SubNumber': int
            ,'SubName': str
            ,'Latitude': float
            ,'SubLatitude': float
            ,'Longitude': float
            ,'SubLongitude': float
            ,'AreaNum': int
            ,'OwnerNumber': int
            ,'OwnerName': str
            ,'DataMaintainer': str
        }
        df = wpp_lib.get_param_df(SimAuto, table, parameter_type)
        df = df[df['SubNumber'].notnull()]

        df['Miles Difference'] = df.apply(
            lambda row: haversine(row['Longitude'], row['Latitude'], row['SubLongitude'], row['SubLatitude']),
            axis=1)
        df = df[df['Miles Difference'] > threshold_miles]
        df['>200kV'] = df['NomkV'].apply(lambda x: 'YES' if x > 200 else 'NO')
        df = df.sort_values(by=['>200kV','AreaNum','NomkV','Number'], ascending=[False, True, False, True])
        return df

    def subnum_not_in_busnums() -> pd.DataFrame:
        """Report Substation numbers that do not exist in the list of that substation's bus numbers."""
        table = 'Bus'
        parameter_type: dict[str,type] = {
            'Number': int
            ,'SubNumber': int
        }
        df_bus = wpp_lib.get_param_df(SimAuto, table, parameter_type)
        df_bus = df_bus.dropna(subset=['SubNumber'])
        # The SubNumber is nullable, and may be nan. 
        # Therefore it likely got converted to a "float" to handle the nullable types. (int is not a nullable type!)
        # Now that it's filtered to only real values, convert it properly to int. 
        df_bus['SubNumber'] = df_bus['SubNumber'].astype(int)
        # Get a dictionary of key=SubNumber to value=list[Number] where bus number is sorted.
        # Example... {640170: [640170, 640171, 640172, 640173, 640174, 640175]}
        sub_bus_dict: dict[int, list[int]] = df_bus.groupby('SubNumber')['Number'].apply(sorted).to_dict()
        sub_bus_df = pd.DataFrame.from_dict(sub_bus_dict, orient='index')
        sub_bus_df.index.name = 'SubNumber'

        subnum_valid = {}
        for subnum, buslist in sub_bus_dict.items():
            subnum_valid[subnum] = 'YES' if subnum in buslist else 'NO'

        table = 'Substation'
        parameter_type: dict[str,type] = {
            'Number': int
            ,'Name': str
            ,'NomkVMax': float
            ,'NomkVMin': float
            ,'AreaNumber': int
            ,'ZoneName': str
            ,'DataMaintainer': str
        }
        df_sub = wpp_lib.get_param_df(SimAuto, table, parameter_type)
        df_sub.loc[:, '>200kV'] = df_sub['NomkVMax'].gt(200).map({True: 'YES', False: 'NO'})
        df_sub['Valid Number'] = df_sub['Number'].map(subnum_valid)

        df = pd.merge(df_sub, sub_bus_df, left_on='Number', right_on='SubNumber', how='left')

        df = df.sort_values(by=['Valid Number','>200kV','AreaNumber'], ascending=[False, False, True])

        df = df[df['Valid Number']=='NO']
        df = df.replace('', pd.NA).dropna(axis=1, how='all') # Drop empty columns

        return df

    def sub_missing_rground() -> pd.DataFrame:
        """Report Substations with missing R ground (Rground==0)."""
        table = 'Substation'
        parameter_type: dict[str,type] = {
            'Number': int
            ,'Name': str
            ,'NomkVMax': float
            ,'NomkVMin': float
            ,'Rground': float
            # ,'Latitude': float
            # ,'Longitude': float
            ,'AreaNumber': int
            ,'ZoneName': str
            ,'DataMaintainer': str
        }
        df = wpp_lib.get_param_df(SimAuto, table, parameter_type)
        df.loc[:, '>200kV'] = df['NomkVMax'].gt(200).map({True: 'YES', False: 'NO'})
        df = df[df['Rground']==0]
        df = df.sort_values(by=['>200kV','AreaNumber'], ascending=[False, True])
        return df

    def transformer_with_length(threshold_miles: float = 0.5) -> pd.DataFrame:
        """Report transformers which have length according to the Lat/Long of the From/To buses."""
        table = 'Branch'
        parameter_type: dict[str,type] = {
            'BusNumFrom': int
            ,'BusNameFrom': str
            ,'NomkVFrom': float
            ,'BusNumTo': int
            ,'BusNameTo': str
            ,'NomkVTo': float
            ,'Circuit': str
            ,'BranchDeviceType': str
            ,'GICLineDistanceMile': float
            ,'AreaNumberFrom': int
            ,'OwnerNum1': int
            ,'OwnerName1': str
            ,'DataMaintainer': str
        }
        df = wpp_lib.get_param_df(SimAuto, table, parameter_type)
        df = df[
            df['BranchDeviceType'].str.contains('Transformer') &
            (df['GICLineDistanceMile'] > threshold_miles)
            ]
        df['>200kV'] = ((df['NomkVFrom'] > 200) | (df['NomkVTo'] > 200)).map({True: 'YES', False: 'NO'})
        df = df.sort_values(by=['>200kV','AreaNumberFrom'], ascending=[False, True])
        return df

    def transformer_missing_data() -> pd.DataFrame:
        """Report Transformers with 'Unknown' GICCoreType, XFConfiguration, or GICAutoXF."""
        table = 'Branch'
        parameter_type: dict[str,type] = {
            'BusNumFrom': int
            ,'BusNameFrom': str
            ,'NomkVFrom': float
            ,'BusNumTo': int
            ,'BusNameTo': str
            ,'NomkVTo': float
            ,'Circuit': str
            ,'BranchDeviceType': str
            ,'GICCoreType': str
            ,'XFConfiguration': str
            ,'GICAutoXF': str
            ,'NomkVMax': float
            ,'AreaNumberFrom': int
            ,'OwnerNum1': int
            ,'OwnerName1': str
            ,'DataMaintainer': str
        }
        df = wpp_lib.get_param_df(SimAuto, table, parameter_type)
        df = df[(df['BranchDeviceType'].str.contains('Transformer')) & 
                (
                    (df['GICCoreType'] == 'Unknown') | 
                    (df['XFConfiguration'] == 'Unknown') | 
                    (df['GICAutoXF'] == 'Unknown')
                )
                ]
        df.loc[:, '>200kV'] = df['NomkVMax'].gt(200).map({True: 'YES', False: 'NO'})
        df = df.sort_values(by=['>200kV','AreaNumberFrom'], ascending=[False, True])
        return df

    def line_length_mismatch(threshold_miles: float = 0.5, threshold_ratio: float = 0.5) -> pd.DataFrame:
        """
            Report lines with significant mismatch between:
            - Length estimated from From/To bus Lat/Long, vs
            - Length estimated from Impedance (R/X/B) parameters.

            Filters on:
            - Absolute difference > `threshold_miles`, and
            - Difference in ratio of the two > `threshold_ratio`
        """
        table = 'Branch'
        parameter_type: dict[str,type] = {
            'BusNumFrom': int
            ,'BusNameFrom': str
            ,'NomkVFrom': float
            ,'BusNumTo': int
            ,'BusNameTo': str
            ,'NomkVTo': float
            ,'Circuit': str
            ,'BranchDeviceType': str
            ,'LineLength': float
            ,'GICLineDistanceMile': float
            ,'LineLengthXBMiles': float
            # ,'LineLengthXBRatio': float # GICLineDistanceMile / LineLengthXBMiles
            ,'AreaNumberFrom': int
            ,'OwnerNum1': int
            ,'OwnerName1': str
            ,'DataMaintainer': str
        }
        df = wpp_lib.get_param_df(SimAuto, table, parameter_type)
        df['LatLong vs XB (Ratio)'] = abs(df['GICLineDistanceMile'] / df['LineLengthXBMiles'] - 1)
        df['LatLong vs XB (Miles)'] = abs(df['GICLineDistanceMile'] - df['LineLengthXBMiles'])
        df = df[
            # Lines only. 
            df['BranchDeviceType'].str.contains('Line') &
            # Length difference is greater than the threshold in miles. 
            (df['LatLong vs XB (Miles)'] > threshold_miles) &
            # The ratio of the two length estimates is more than the threshold ratio. 
            (df['LatLong vs XB (Ratio)'] > threshold_ratio)
            ]
        df['>200kV'] = ((df['NomkVFrom'] > 200) | (df['NomkVTo'] > 200)).map({True: 'YES', False: 'NO'})
        df = df.sort_values(by=['>200kV','AreaNumberFrom','LatLong vs XB (Miles)'], ascending=[False, True, False])
        return df

    def line_r_suspect(acdc_high: float = 1.2, acdc_low: float = 0.99) -> pd.DataFrame:
        """
        Report Lines with high mismatch between PowerFlow R1 and custom entered DC resistance.
        - DC Resistance should be lower than Positive Sequence AC resistance (R1). 
        - DC Resistance should be close to Positive Sequence AC resistance (R1) (default threshod: within ~20%)
        """
        table = 'Branch'
        parameter_type: dict[str,type] = {
            'BusNumFrom': int
            ,'BusNameFrom': str
            ,'NomkVFrom': float
            ,'BusNumTo': int
            ,'BusNameTo': str
            ,'NomkVTo': float
            ,'Circuit': str
            ,'BranchDeviceType': str
            ,'GICUSEPFR': str # YES or NO, will we be using a custom DC R value?
            ,'GICCUSTOMR1': float # Custom R entry.
            ,'GICPFR1': float # DC R estimated from PowerFlow R1. 
            # GICR is the resistance which is used for the simulation. 
            # It is a calculated field, not user-entered. 
            #   If GICUSEPFR="YES", this value comes from GICCUSTOMR1. 
            #   If GICUSEPFR="NO", this value comes from GICPFR1.
            # ,'GICR': float 
            ,'AreaNumberFrom': int
            ,'OwnerNum1': int
            ,'OwnerName1': str
            ,'DataMaintainer': str
        }
        df = wpp_lib.get_param_df(SimAuto, table, parameter_type)
        df['>200kV'] = ((df['NomkVFrom'] > 200) | (df['NomkVTo'] > 200)).map({True: 'YES', False: 'NO'})
        df = df[
            # Lines only. 
            df['BranchDeviceType'].str.contains('Line') &
            # Filter by lines where the user specified that the custom resistance must be used. 
            (df['GICUSEPFR'] == 'YES') & 
            (
                # User-specified DC Resistance must be lower than AC resistance. Using a 1% tolerance to avoid false-flagging due to precision errors. 
                ((df['GICPFR1'] / df['GICCUSTOMR1']) < acdc_low) | 
                # DC Resistance is rarely less than R1 by 20% or more. 
                ((df['GICPFR1'] / df['GICCUSTOMR1']) > acdc_high)
            )
            ]
        df = df.sort_values(by=['>200kV','AreaNumberFrom'], ascending=[False, True])
        return df

    def line_from_to_kv_difference() -> pd.DataFrame:
        """Report Lines with a mismatch in nominal kV between the From/To sides."""
        table = 'Branch'
        parameter_type: dict[str,type] = {
            'BusNumFrom': int
            ,'BusNameFrom': str
            ,'NomkVFrom': float
            ,'BusNumTo': int
            ,'BusNameTo': str
            ,'NomkVTo': float
            ,'Circuit': str
            ,'BranchDeviceType': str
            ,'AreaNumberFrom': int
            ,'OwnerNum1': int
            ,'OwnerName1': str
            ,'DataMaintainer': str
        }
        df = wpp_lib.get_param_df(SimAuto, table, parameter_type)
        df['>200kV'] = ((df['NomkVFrom'] > 200) | (df['NomkVTo'] > 200)).map({True: 'YES', False: 'NO'})
        df = df[df['BranchDeviceType'].str.contains('Line')]
        df = df[df['NomkVFrom'] != df['NomkVTo']]
        df = df.sort_values(by=['>200kV','AreaNumberFrom'], ascending=[False, True])
        return df

    def sub_area_ne_bus_area() -> pd.DataFrame:
        """
        First, estimates substation area number based on the buses which it contains. 
        If any buses don't match the substation area, it will be reported. 
        """
        table = 'Substation'
        parameter_type: dict[str,type] = {
            'Number': int
            ,'Name': str
            ,'AreaNumber': int
            ,'DataMaintainer': str
        }
        df_sub = wpp_lib.get_param_df(SimAuto, table, parameter_type)

        table = 'Bus'
        parameter_type: dict[str,type] = {
            'Number': int
            ,'Name': str
            ,'NomkV': float
            ,'AreaNumber': int
            ,'ZoneName': str
            ,'DataMaintainer': str
            ,'SubNumber': int
        }
        df_bus = wpp_lib.get_param_df(SimAuto, table, parameter_type)

        # Join df_bus SubNumber with df_sub Number. Use 'Bus' suffix and 'Sub' prefix. 
        df = df_bus.merge(df_sub, left_on='SubNumber', right_on='Number', how='outer', suffixes=('_Bus', '_Sub'))
        # Filter on AreaNumber_Bus != AreaNumber_Sub.
        df = df[df['AreaNumber_Bus'] != df['AreaNumber_Sub']]
        # Drop NA in the Number_Sub column.
        df = df.dropna(subset=['Number_Sub'])

        df.loc[:, '>200kV'] = df['NomkV'].gt(200).map({True: 'YES', False: 'NO'})
        df = df.sort_values(by=['>200kV','AreaNumber_Bus'], ascending=[False, True])
        return df

    if not wpp_lib.open_case(SimAuto, pw_fp):
        return
    
    # Get reports.
    report_dict: dict[str, pd.DataFrame] = {
        'Bus - Undefined Sub': bus_unmapped_sub()
        ,'Bus vs Sub - LatLong': bus_sub_latlong_mismatch(threshold_miles=0.5)
        ,'Bus vs Sub - Area': sub_area_ne_bus_area()
        ,'SubNum not in BusNums': subnum_not_in_busnums()
        ,'Sub Missing Rground': sub_missing_rground()
        ,'XFMR with Length': transformer_with_length(threshold_miles=0.5)
        ,'XFMR Missing Data': transformer_missing_data()
        ,'Line Length Suspect': line_length_mismatch(threshold_miles=0.5, threshold_ratio=0.5)
        ,'Line R Suspect': line_r_suspect(acdc_high=1.2, acdc_low=0.99)
        ,'Line Changes NomkV': line_from_to_kv_difference()
    }

    return report_dict

if(__name__=='__main__'):
    # Open PowerWorld.
    SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

    # Run Quality Checks.
    report_dict = run_gmd_quality_checks(SimAuto, pw_fp, rep_fp)

    # Save reports.
    wpp_lib.df_dict_to_excel_workbook(rep_fp, report_dict)

    # Close PowerWorld. 
    SimAuto.CloseCase()
    SimAuto = None
    print('done')

