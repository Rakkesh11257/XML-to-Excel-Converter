import os
import xml.etree.ElementTree as ET
import pandas as pd

# Function to extract data from XML files and create Excel
def create_excel_from_xml(folder_path, output_excel):
    # Create empty DataFrames to store archived, deleted, and approved assets with missing data
    archived_assets_list = []
    deleted_assets_list = []
    approved_assets_missing_data = []

    # Iterate through XML files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.xml'):
            file_path = os.path.join(folder_path, filename)
            try:
                tree = ET.parse(file_path)
                root = tree.getroot()

                # Find all DownstreamsModel tags
                downstreams_models = root.findall('.//DownstreamsModel')

                for downstream_model in downstreams_models:
                    digital_asset_media_attributes = downstream_model.find('.//DigitalAssetMediaAttributes')
                    digital_asset_id = downstream_model.find('.//DigitalAssetID').text
                    digital_asset_final_lifecycle_status = downstream_model.find('.//DigitalAssetFinalLifeCycleStatus').text

                    # Check if DigitalAssetFinalLifeCycleStatus is "Archived"
                    if digital_asset_final_lifecycle_status.lower() == 'archived':
                        # Append archived asset to the list with file name
                        archived_assets_list.append({'Asset ID': digital_asset_id, 'File Name': filename})
                    # Check if DigitalAssetFinalLifeCycleStatus is "Deleted"
                    elif digital_asset_final_lifecycle_status.lower() == 'deleted':
                        # Append deleted asset to the list with file name
                        deleted_assets_list.append({'Asset ID': digital_asset_id, 'File Name': filename})
                    # Check if DigitalAssetFinalLifeCycleStatus is "Approved"
                    elif digital_asset_final_lifecycle_status.lower() == 'approved':
                        # Check for missing data in specific tags within DigitalAssetMediaAttributes
                        missing_data = []
                        digital_asset_view = downstream_model.find('.//DigitalAssetView')
                        digital_asset_angle = downstream_model.find('.//DigitalAssetAngle')
                        digital_asset_public_link_jpg_elem = downstream_model.find('.//DigitalAssetPublicLinkRenditionURLJPG')
                        digital_asset_public_link_png_elem = downstream_model.find('.//DigitalAssetPublicLinkRenditionURLPNG')

                        # Check if any of the tags are empty (i.e., have no value)
                        if not all([digital_asset_view is not None and digital_asset_view.text,
                                    digital_asset_angle is not None and digital_asset_angle.text,
                                    digital_asset_public_link_jpg_elem is not None and digital_asset_public_link_jpg_elem.text,
                                    digital_asset_public_link_png_elem is not None and digital_asset_public_link_png_elem.text]):
                            if digital_asset_view is None or not digital_asset_view.text:
                                missing_data.append('DigitalAssetView')
                            if digital_asset_angle is None or not digital_asset_angle.text:
                                missing_data.append('DigitalAssetAngle')
                            if digital_asset_public_link_jpg_elem is None or not digital_asset_public_link_jpg_elem.text:
                                missing_data.append('DigitalAssetPublicLinkRenditionURLJPG')
                            if digital_asset_public_link_png_elem is None or not digital_asset_public_link_png_elem.text:
                                missing_data.append('DigitalAssetPublicLinkRenditionURLPNG')

                            # Append asset with missing data to the list along with the file name
                            approved_assets_missing_data.append({'Asset ID': digital_asset_id, 'Missing Data': ', '.join(missing_data), 'File Name': filename})
            except ET.ParseError as e:
                print(f"Error parsing {filename}: {e}")
                continue

    # Create DataFrames from the lists of archived, deleted, and approved assets with missing data
    archived_assets = pd.DataFrame(archived_assets_list)
    deleted_assets = pd.DataFrame(deleted_assets_list)
    approved_assets_missing = pd.DataFrame(approved_assets_missing_data)

    # Write archived, deleted, and approved assets with missing data to Excel
    with pd.ExcelWriter(output_excel) as writer:
        archived_assets.to_excel(writer, sheet_name='Archived Assets', index=False)
        deleted_assets.to_excel(writer, sheet_name='Deleted Assets', index=False)
        approved_assets_missing.to_excel(writer, sheet_name='Approved Assets with Missing Data', index=False)

# Example usage
folder_path = r"C:\Users\rakkesh_r\Documents\Pathfinder XML files\2nd April"
output_excel = r'C:\Users\rakkesh_r\Desktop\output.xlsx'
create_excel_from_xml(folder_path, output_excel)
