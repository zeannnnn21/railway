"""
Report generation functions
Adapted from existing Python scripts to work with JSON input
"""

import tempfile
import os
from pathlib import Path
import pandas as pd
from datetime import datetime
import io

def json_to_qualys_csv(data):
    """
    Convert JSON scan data to Qualys CSV format in memory
    
    Args:
        data: Dict with keys: scan, vulnerabilities, assets
    
    Returns:
        str: CSV content in Qualys format
    """
    scan = data['scan']
    vulnerabilities = data.get('vulnerabilities', [])
    assets = data.get('assets', [])
    
    lines = []
    
    # Lines 1-3: Report metadata
    lines.append('"Report Type","Scan"')
    lines.append(f'"Generated","{datetime.now().isoformat()}"')
    lines.append('""')
    
    # Lines 4-5: Empty
    lines.append('')
    lines.append('')
    
    # Line 6: Scan metadata
    scan_date = scan.get('scan_date', datetime.now().isoformat())
    active_hosts = scan.get('active_hosts', 0)
    total_hosts = scan.get('total_hosts', 0)
    
    metadata_headers = [
        'Launch Date',
        'Active Hosts',
        'Total Hosts',
        'Not Scanned',
        'Duration',
        'Scan Title',
        'Asset Groups',
        'IPs',
        'Scan Reference'
    ]
    
    metadata_values = [
        scan_date,
        str(active_hosts),
        str(total_hosts),
        str(total_hosts - active_hosts),
        '',
        scan.get('scan_title', ''),
        scan.get('client_business_unit', ''),
        '',
        scan.get('document_control_number', '')
    ]
    
    lines.append('"' + '","'.join(metadata_headers) + '"')
    lines.append('"' + '","'.join(metadata_values) + '"')
    
    # Line 7: Empty
    lines.append('')
    
    # Line 8: Vulnerability headers
    vuln_headers = [
        'IP', 'DNS', 'NetBIOS', 'OS', 'IP Status', 'QID', 'Title', 'Type', 'Severity',
        'Port', 'Protocol', 'FQDN', 'SSL', 'CVE ID', 'Vendor Reference', 'Bugtraq ID',
        'Threat', 'Impact', 'Solution', 'Exploitability', 'Associated Malware', 'Results',
        'PCI Vuln', 'Instance', 'Category', 'Associated Tags', 'CVSSv3 Base',
        'CVSSv3 Temporal', 'CVSS Base', 'CVSS Temporal', 'First Detected',
        'Last Detected', 'Times Detected', 'Date Updated', 'Patchable'
    ]
    lines.append('"' + '","'.join(vuln_headers) + '"')
    
    # Lines 9+: Vulnerabilities
    severity_map = {
        'low': '1',
        'medium': '2', 
        'high': '3',
        'critical': '5'
    }
    
    for vuln in vulnerabilities:
        # Get asset info (vulnerabilities should have asset_id reference)
        asset = vuln.get('asset', {})
        
        severity = severity_map.get(vuln.get('severity', 'low').lower(), '1')
        
        row = [
            asset.get('ip_address', ''),
            asset.get('hostname', ''),
            '',  # NetBIOS
            asset.get('operating_system', ''),
            'Active',
            vuln.get('plugin_id', ''),
            vuln.get('vulnerability_name', ''),
            'Vuln',
            severity,
            vuln.get('port', ''),
            vuln.get('protocol', ''),
            '',  # FQDN
            '',  # SSL
            vuln.get('cve_ids', ''),
            '',  # Vendor Reference
            '',  # Bugtraq ID
            vuln.get('threat_description', ''),
            vuln.get('impact_description', ''),
            vuln.get('solution_description', ''),
            '',  # Exploitability
            '',  # Associated Malware
            vuln.get('vulnerability_proof', ''),
            '',  # PCI Vuln
            '',  # Instance
            '',  # Category
            '',  # Associated Tags
            '',  # CVSSv3 Base
            '',  # CVSSv3 Temporal
            vuln.get('cvss_score', ''),
            '',  # CVSS Temporal
            vuln.get('first_detected', ''),
            vuln.get('last_detected', ''),
            str(vuln.get('times_detected', 1)),
            '',  # Date Updated
            ''   # Patchable
        ]
        
        # Escape and quote
        escaped_row = [f'"{str(field).replace(chr(34), chr(34)+chr(34))}"' for field in row]
        lines.append(','.join(escaped_row))
    
    return '\n'.join(lines)


def generate_detailed_report(data):
    """
    Generate detailed Excel report from JSON data
    
    Args:
        data: Dict with scan data
    
    Returns:
        str: Path to generated Excel file
    """
    # Convert JSON to CSV
    csv_content = json_to_qualys_csv(data)
    
    # Write CSV to temp file
    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as csv_file:
        csv_file.write(csv_content)
        csv_path = csv_file.name
    
    # Get template path
    template_path = Path(__file__).parent / 'templates' / 'Template.xlsx'
    
    # Create output file
    output_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    output_path = output_file.name
    output_file.close()
    
    # Import and run the existing detailed report script logic
    # For now, using a simplified version - you'll need to adapt the full script
    import sys
    sys.path.append(str(Path(__file__).parent / 'scripts'))
    
    try:
        from qualys_report_automation_v3_2 import QualysReportAutomation
        
        automation = QualysReportAutomation()
        business_unit = data['scan'].get('client_business_unit', 'Unknown')
        
        success = automation.fill_vulnerability_details(
            str(template_path),
            csv_path,
            output_path,
            business_unit
        )
        
        if not success:
            raise Exception("Report generation failed")
        
        # Cleanup temp CSV
        os.unlink(csv_path)
        
        return output_path
        
    except Exception as e:
        # Cleanup on error
        if os.path.exists(csv_path):
            os.unlink(csv_path)
        if os.path.exists(output_path):
            os.unlink(output_path)
        raise


def generate_executive_report(data):
    """
    Generate executive Word report from JSON data
    
    Args:
        data: Dict with scan data
    
    Returns:
        str: Path to generated Word file
    """
    # Convert JSON to CSV
    csv_content = json_to_qualys_csv(data)
    
    # Write CSV to temp file
    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as csv_file:
        csv_file.write(csv_content)
        csv_path = csv_file.name
    
    # Get template path
    template_path = Path(__file__).parent / 'templates' / 'Report-Template-with-Chart.docx'
    
    # Create output file
    output_file = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
    output_path = output_file.name
    output_file.close()
    
    # Import and run the existing executive report script logic
    import sys
    sys.path.append(str(Path(__file__).parent / 'scripts'))
    
    try:
        from executive_report_automation_v1_5 import ExecutiveReportAutomation
        
        automation = ExecutiveReportAutomation()
        client_name = data.get('clientName', data['scan'].get('client_business_unit', 'Unknown Client'))
        
        success = automation.generate_report(
            csv_path,
            str(template_path),
            output_path,
            client_name
        )
        
        if not success:
            raise Exception("Report generation failed")
        
        # Cleanup temp CSV
        os.unlink(csv_path)
        
        return output_path
        
    except Exception as e:
        # Cleanup on error
        if os.path.exists(csv_path):
            os.unlink(csv_path)
        if os.path.exists(output_path):
            os.unlink(output_path)
        raise
