"""
Report generation functions
Calls existing Python scripts via subprocess
"""

import tempfile
import os
from pathlib import Path
from datetime import datetime
import subprocess

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
    
    return '\n'.join(lines)  # FIXED: single backslash for actual newline


def generate_detailed_report(data):
    """
    Generate detailed Excel report from JSON data using the Python script
    
    Args:
        data: Dict with scan data
    
    Returns:
        str: Path to generated Excel file
    """
    print(f"Generating detailed report for scan: {data['scan'].get('document_control_number', 'unknown')}")
    print("=" * 60)
    
    # Convert JSON to CSV
    csv_content = json_to_qualys_csv(data)
    
    # Write CSV to temp file
    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as csv_file:
        csv_file.write(csv_content)
        csv_path = csv_file.name
    
    print(f"Processing: {csv_path}")
    print("=" * 60)
    
    # Get paths
    script_path = Path(__file__).parent / 'scripts' / 'qualys_report_automation_v3_2.py'
    template_path = Path(__file__).parent / 'templates' / 'Template.xlsx'
    
    # Create output file
    output_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    output_path = output_file.name
    output_file.close()
    
    business_unit = data.get('businessUnit', data['scan'].get('client_business_unit', 'Unknown'))
    
    try:
        # Call Python script directly
        result = subprocess.run(
            ['python', str(script_path), csv_path, str(template_path), output_path, business_unit],
            capture_output=True,
            text=True,
            timeout=180
        )
        
        # Print script output
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print(result.stderr)
        
        # Check if output file exists (more reliable than exit code)
        if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
            raise Exception(f"Output file not created or empty. Exit code: {result.returncode}")
        
        print(f"✓ Report successfully generated: {output_path}")
        
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
    Generate executive Word report from JSON data using the Python script
    
    Args:
        data: Dict with scan data
    
    Returns:
        str: Path to generated Word file
    """
    print(f"Generating executive report for scan: {data['scan'].get('document_control_number', 'unknown')}")
    print("=" * 60)
    
    # Convert JSON to CSV
    csv_content = json_to_qualys_csv(data)
    
    # Write CSV to temp file
    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8') as csv_file:
        csv_file.write(csv_content)
        csv_path = csv_file.name
    
    print(f"Processing: {csv_path}")
    print("=" * 60)
    
    # Get paths
    script_path = Path(__file__).parent / 'scripts' / 'executive_report_automation_v1_5.py'
    template_path = Path(__file__).parent / 'templates' / 'Report-Template-with-Chart.docx'
    
    # Create output file
    output_file = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
    output_path = output_file.name
    output_file.close()
    
    client_name = data.get('clientName', data['scan'].get('client_business_unit', 'Unknown Client'))
    
    try:
        # Call Python script directly
        result = subprocess.run(
            ['python', str(script_path), csv_path, str(template_path), output_path, client_name],
            capture_output=True,
            text=True,
            timeout=180
        )
        
        # Print script output
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print(result.stderr)
        
        # Check if output file exists (more reliable than exit code)
        if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
            raise Exception(f"Output file not created or empty. Exit code: {result.returncode}")
        
        print(f"✓ Report successfully generated: {output_path}")
        
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
