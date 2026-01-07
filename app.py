"""
VM Monitoring Report Generation Service
Flask API for generating Excel and Word reports
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import tempfile
import os
import json
from pathlib import Path
import sys

# Add current directory to path
sys.path.append(os.path.dirname(__file__))

from report_generators import generate_detailed_report, generate_executive_report

app = Flask(__name__)
CORS(app)  # Enable CORS for Vercel requests

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        'service': 'VM Monitoring Report Generator',
        'version': '1.0.0',
        'status': 'healthy',
        'endpoints': {
            'detailed': '/api/reports/detailed',
            'executive': '/api/reports/executive',
            'health': '/health'
        }
    })

@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy',
        'message': 'Report generation service is running'
    })

@app.route('/api/reports/detailed', methods=['POST'])
def detailed_report():
    """
    Generate detailed Excel report from scan data
    
    Expected JSON payload:
    {
        "scan": { ... scan metadata ... },
        "vulnerabilities": [ ... vulnerability list ... ],
        "assets": [ ... asset list ... ]
    }
    """
    try:
        data = request.json
        
        if not data:
            return jsonify({'error': 'No data provided'}), 400
        
        # Validate required fields
        if 'scan' not in data or 'vulnerabilities' not in data:
            return jsonify({'error': 'Missing required fields: scan, vulnerabilities'}), 400
        
        # Generate report
        print(f"Generating detailed report for scan: {data['scan'].get('document_control_number', 'unknown')}")
        report_path = generate_detailed_report(data)
        
        # Read file into memory and cleanup immediately
        try:
            with open(report_path, 'rb') as f:
                file_data = f.read()
            os.unlink(report_path)  # Delete temp file immediately
        except Exception as e:
            if os.path.exists(report_path):
                os.unlink(report_path)
            raise
        
        # Send file from memory
        from flask import Response
        response = Response(
            file_data,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response.headers['Content-Disposition'] = f'attachment; filename="Detailed_Report_{data["scan"].get("document_control_number", "report")}.xlsx"'
        return response
        
    except Exception as e:
        print(f"Error generating detailed report: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({
            'error': 'Report generation failed',
            'details': str(e)
        }), 500

@app.route('/api/reports/executive', methods=['POST'])
def executive_report():
    """
    Generate executive Word report from scan data
    
    Expected JSON payload:
    {
        "scan": { ... scan metadata ... },
        "vulnerabilities": [ ... vulnerability list ... ],
        "assets": [ ... asset list ... ],
        "clientName": "Client Name"
    }
    """
    try:
        data = request.json
        
        if not data:
            return jsonify({'error': 'No data provided'}), 400
        
        # Validate required fields
        if 'scan' not in data or 'vulnerabilities' not in data:
            return jsonify({'error': 'Missing required fields: scan, vulnerabilities'}), 400
        
        # Generate report
        print(f"Generating executive report for scan: {data['scan'].get('document_control_number', 'unknown')}")
        report_path = generate_executive_report(data)
        
        # Read file into memory and cleanup immediately
        try:
            with open(report_path, 'rb') as f:
                file_data = f.read()
            os.unlink(report_path)  # Delete temp file immediately
        except Exception as e:
            if os.path.exists(report_path):
                os.unlink(report_path)
            raise
        
        # Send file from memory
        from flask import Response
        response = Response(
            file_data,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        response.headers['Content-Disposition'] = f'attachment; filename="Executive_Report_{data["scan"].get("document_control_number", "report")}.docx"'
        return response
        
    except Exception as e:
        print(f"Error generating executive report: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({
            'error': 'Report generation failed',
            'details': str(e)
        }), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
