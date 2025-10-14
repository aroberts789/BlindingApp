#!/usr/bin/env python3
"""
File Blinder Local Web Server
A complete executable that runs a web server for file processing
with persistent keyword management
"""

import os
import sys
import json
import webbrowser
import threading
import tempfile
import traceback
from pathlib import Path
from flask import Flask, render_template_string, request, jsonify, send_file
from werkzeug.utils import secure_filename

# Add the current directory to Python path so we can import file_blinder
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

try:
    from file_blinder import FileBlinder

    FILEBLINDER_AVAILABLE = True
except ImportError:
    FILEBLINDER_AVAILABLE = False
    print("Warning: file_blinder.py not found. Please ensure it's in the same directory.")

# Flask app configuration
app = Flask(__name__)
app.secret_key = 'file-blinder-secret-key-2024'
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size

# Allowed file extensions
ALLOWED_EXTENSIONS = {'.docx', '.html', '.htm', '.txt'}

# Keywords storage file
KEYWORDS_FILE = current_dir / 'keywords.json'


def allowed_file(filename):
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def load_keywords():
    """Load keywords from persistent storage"""
    default_keywords = [
        {"original": "confidential", "replacement": "REDACTED", "enabled": True},
        {"original": "secret", "replacement": "REDACTED", "enabled": True},
        {"original": "internal", "replacement": "INTERNAL", "enabled": True},
        {"original": "proprietary", "replacement": "PROPRIETARY", "enabled": True},
        {"original": "classified", "replacement": "CLASSIFIED", "enabled": True},
        {"original": r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', "replacement": "EMAIL", "enabled": True},
        {"original": r'\b\d{3}[-.]?\d{3}[-.]?\d{4}\b', "replacement": "PHONE", "enabled": True},
        {"original": r'\b\d{3}-\d{2}-\d{4}\b', "replacement": "SSN", "enabled": True},
        {"original": r'\b\d{1,2}/\d{1,2}/\d{4}\b', "replacement": "DATE", "enabled": True},
        {"original": r'\$[\d,]+\.?\d*', "replacement": "AMOUNT", "enabled": True}
    ]

    try:
        if KEYWORDS_FILE.exists():
            with open(KEYWORDS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('keywords', default_keywords)
        else:
            save_keywords(default_keywords)
            return default_keywords
    except Exception as e:
        print(f"Error loading keywords: {e}")
        return default_keywords


def save_keywords(keywords):
    """Save keywords to persistent storage"""
    try:
        data = {
            'keywords': keywords,
            'version': '1.0',
            'last_updated': str(Path(__file__).stat().st_mtime)
        }
        with open(KEYWORDS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving keywords: {e}")
        return False


def get_active_keywords():
    """Get only enabled keywords as a dictionary for FileBlinder"""
    keywords = load_keywords()
    return {kw['original']: kw['replacement'] for kw in keywords if kw.get('enabled', True)}


@app.route('/preview', methods=['POST'])
def preview_file():
    """Generate preview showing original vs processed document"""
    if not FILEBLINDER_AVAILABLE:
        return jsonify({'error': 'FileBlinder module not available'}), 500

    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        if not allowed_file(file.filename):
            return jsonify({'error': f'File type not supported'}), 400

        # Secure the filename
        filename = secure_filename(file.filename)

        # Get form parameters
        processing_method = request.form.get('processing_method', 'safe')
        standardize_formatting = request.form.get('standardize_formatting') == 'on'
        font_name = request.form.get('font_name', 'Calibri')
        try:
            font_size = int(request.form.get('font_size', 11))
        except ValueError:
            font_size = 11
        font_color_black = request.form.get('font_color_black') == 'on'

        # Create temporary files
        with tempfile.NamedTemporaryFile(delete=False, suffix=Path(filename).suffix) as tmp_input:
            file.save(tmp_input.name)
            input_path = tmp_input.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=Path(filename).suffix) as tmp_output:
            output_path = tmp_output.name

        try:
            # Get active keywords
            active_keywords = get_active_keywords()

            # Initialize FileBlinder
            blinder = FileBlinder(
                keyword_replacements=active_keywords,
                standardize_formatting=standardize_formatting,
                font_name=font_name,
                font_size=font_size,
                font_color_black=font_color_black,
                grey_shading=False
            )

            # Extract original structure
            print("Extracting original document structure...")
            original_structure = blinder.extract_document_structure(input_path)

            # Process the file
            print("Processing file...")
            blinder.blind_file(
                input_path=input_path,
                output_path=output_path,
                method=processing_method
            )

            # Extract processed structure
            print("Extracting processed document structure...")
            processed_structure = blinder.extract_document_structure(output_path)

            # Generate diff
            print("Generating diff...")
            diff_data = blinder.generate_diff(original_structure, processed_structure)

            # Calculate statistics
            stats = {
                'images_removed': len(diff_data['image_changes']),
                'keywords_replaced': sum(len(pc['changes']) for pc in diff_data['paragraph_changes']),
                'formatting_changes': len(diff_data['formatting_changes']),
                'table_shadings_removed': len(diff_data['table_changes'])
            }

            # Return preview data
            return jsonify({
                'filename': filename,
                'original': original_structure,
                'processed': processed_structure,
                'diff': diff_data,
                'statistics': stats
            })

        finally:
            # Clean up temporary files
            try:
                os.unlink(input_path)
                os.unlink(output_path)
            except:
                pass

    except Exception as e:
        print(f"Error generating preview: {e}")
        print(traceback.format_exc())
        return jsonify({'error': f'Preview generation error: {str(e)}'}), 500
# Complete HTML template with keyword management
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>File Blinder - Document Anonymization Tool</title>
    <link rel="icon" href="data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><text y=%22.9em%22 font-size=%2290%22>🔒</text></svg>">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: linear-gradient(135deg, #9f237e 0%, #313e4B 100%);
            min-height: 100vh;
            color: #333;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }

        .header {
            text-align: center;
            color: white;
            margin-bottom: 30px;
        }

        .header h1 {
            font-size: 2.5rem;
            margin-bottom: 10px;
            margin-top: 0;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }

        .header p {
            font-size: 1.1rem;
            opacity: 0.9;
        }

        .main-content {
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }

        .tabs {
            display: flex;
            background: #f8f9fa;
            border-bottom: 1px solid #dee2e6;
        }

        .tab {
            flex: 1;
            padding: 20px;
            text-align: center;
            background: none;
            border: none;
            cursor: pointer;
            font-size: 1rem;
            font-weight: 600;
            color: #6c757d;
            transition: all 0.3s ease;
        }

        .tab.active {
            background: white;
            color: #495057;
            border-bottom: 3px solid #667eea;
        }

        .tab:hover {
            background: #e9ecef;
        }

        .tab-content {
            display: none;
            padding: 40px;
        }

        .tab-content.active {
            display: block;
        }

        .upload-area {
            border: 3px dashed #ccc;
            border-radius: 15px;
            padding: 60px 20px;
            text-align: center;
            transition: all 0.3s ease;
            cursor: pointer;
            margin-bottom: 30px;
            position: relative;
        }

        .upload-area.drag-over {
            border-color: #667eea;
            background-color: #f8f9ff;
            transform: scale(1.02);
        }

        .upload-area:hover {
            border-color: #667eea;
            background-color: #f8f9ff;
        }

        .upload-icon {
            font-size: 4rem;
            color: #ccc;
            margin-bottom: 20px;
            transition: all 0.3s ease;
        }

        .upload-area.drag-over .upload-icon,
        .upload-area:hover .upload-icon {
            color: #667eea;
            transform: scale(1.1);
        }

        .upload-text {
            font-size: 1.2rem;
            color: #666;
            margin-bottom: 10px;
        }

        .upload-subtext {
            font-size: 1rem;
            color: #999;
        }

        .file-input {
            display: none;
        }

        .btn {
            background: linear-gradient(45deg, #667eea, #764ba2);
            color: white;
            border: none;
            padding: 15px 35px;
            border-radius: 30px;
            cursor: pointer;
            font-size: 1.1rem;
            font-weight: 600;
            transition: all 0.3s ease;
            text-transform: uppercase;
            letter-spacing: 1px;
            box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
        }

        .btn:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(102, 126, 234, 0.4);
        }

        .btn:disabled {
            background: linear-gradient(45deg, #ccc, #ddd);
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }

        .btn-secondary {
            background: #6c757d;
            box-shadow: 0 4px 15px rgba(108, 117, 125, 0.3);
        }

        .btn-secondary:hover {
            box-shadow: 0 8px 25px rgba(108, 117, 125, 0.4);
        }

        .btn-danger {
            background: #dc3545;
            box-shadow: 0 4px 15px rgba(220, 53, 69, 0.3);
        }

        .btn-danger:hover {
            box-shadow: 0 8px 25px rgba(220, 53, 69, 0.4);
        }

        .btn-success {
            background: #28a745;
            box-shadow: 0 4px 15px rgba(40, 167, 69, 0.3);
        }

        .btn-success:hover {
            box-shadow: 0 8px 25px rgba(40, 167, 69, 0.4);
        }

        .btn-small {
            padding: 8px 16px;
            font-size: 0.9rem;
            border-radius: 20px;
        }

        .file-info {
            background: #f8f9fa;
            border-left: 4px solid #667eea;
            border-radius: 10px;
            padding: 20px;
            margin: 20px 0;
            display: none;
            animation: slideIn 0.3s ease;
        }

        .file-info.show {
            display: block;
        }

        @keyframes slideIn {
            from { opacity: 0; transform: translateY(-10px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .file-info h3 {
            color: #495057;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .settings-section {
            background: #f8f9fa;
            border-radius: 15px;
            padding: 25px;
            margin: 25px 0;
            border: 1px solid #e9ecef;
        }

        .settings-section h3 {
            color: #495057;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .settings-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
        }

        .form-group {
            margin-bottom: 15px;
        }

        .form-group label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #495057;
        }

        .form-group input,
        .form-group select {
            width: 100%;
            padding: 12px;
            border: 2px solid #dee2e6;
            border-radius: 8px;
            font-size: 1rem;
            transition: border-color 0.3s ease;
            background: white;
        }

        .form-group input:focus,
        .form-group select:focus {
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }

        .checkbox-group {
            display: flex;
            align-items: center;
            gap: 10px;
            margin: 15px 0;
        }

        .checkbox-group input[type="checkbox"] {
            width: auto;
            transform: scale(1.2);
        }

        .progress {
            background: #e9ecef;
            border-radius: 10px;
            height: 25px;
            margin: 25px 0;
            overflow: hidden;
            display: none;
            box-shadow: inset 0 2px 4px rgba(0,0,0,0.1);
        }

        .progress.show {
            display: block;
            animation: slideIn 0.3s ease;
        }

        .progress-bar {
            background: linear-gradient(45deg, #667eea, #764ba2);
            height: 100%;
            width: 0%;
            transition: width 0.6s ease;
            border-radius: 10px;
            position: relative;
        }

        .progress-bar::after {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            bottom: 0;
            right: 0;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.3), transparent);
            animation: shimmer 2s infinite;
        }

        @keyframes shimmer {
            0% { transform: translateX(-100%); }
            100% { transform: translateX(100%); }
        }

        .results {
            border-radius: 10px;
            padding: 20px;
            margin: 20px 0;
            display: none;
            animation: slideIn 0.3s ease;
        }

        .results.show {
            display: block;
        }

        .results.success {
            background: #d4edda;
            border: 1px solid #c3e6cb;
            color: #155724;
        }

        .results.error {
            background: #f8d7da;
            border: 1px solid #f5c6cb;
            color: #721c24;
        }

        .supported-formats {
            display: flex;
            justify-content: center;
            gap: 15px;
            margin: 20px 0;
            flex-wrap: wrap;
        }

        .format-badge {
            background: #313e4B;
            color: white;
            padding: 8px 16px;
            border-radius: 20px;
            font-weight: 600;
            font-size: 0.9rem;
            box-shadow: 0 2px 8px rgba(102, 126, 234, 0.3);
        }

        .processing-status {
            text-align: center;
            color: #667eea;
            font-size: 1.2rem;
            font-weight: 600;
            margin: 25px 0;
            display: none;
        }

        .processing-status.show {
            display: block;
            animation: pulse 2s infinite;
        }

        @keyframes pulse {
            0% { opacity: 0.7; }
            50% { opacity: 1; }
            100% { opacity: 0.7; }
        }

        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-top: 15px;
        }

        .info-item {
            background: white;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        .info-item strong {
            color: #667eea;
        }

        .status-bar {
            background: #667eea;
            color: white;
            padding: 10px 20px;
            margin: -40px -40px 30px -40px;
            font-weight: 600;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .version {
            font-size: 0.9rem;
            opacity: 0.8;
        }

        /* Keyword Management Styles */
        .keywords-container {
            max-height: 600px;
            overflow-y: auto;
            border: 1px solid #dee2e6;
            border-radius: 10px;
            background: white;
        }

        .keyword-header {
            background: #f8f9fa;
            padding: 15px 20px;
            border-bottom: 1px solid #dee2e6;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 10;
        }

        .keyword-list {
            padding: 0;
        }

        .keyword-item {
            padding: 15px 20px;
            border-bottom: 1px solid #f1f3f4;
            display: grid;
            grid-template-columns: auto 1fr 1fr auto auto;
            gap: 15px;
            align-items: center;
            transition: background-color 0.2s ease;
        }

        .keyword-item:hover {
            background-color: #f8f9fa;
        }

        .keyword-item:last-child {
            border-bottom: none;
        }

        .keyword-toggle {
            transform: scale(1.2);
        }

        .keyword-input {
            padding: 8px 12px;
            border: 1px solid #dee2e6;
            border-radius: 6px;
            font-size: 0.9rem;
        }

        .keyword-input:focus {
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 0 2px rgba(102, 126, 234, 0.1);
        }

        .keyword-actions {
            display: flex;
            gap: 8px;
        }

        .new-keyword-form {
            padding: 20px;
            background: #f8f9fa;
            border-top: 1px solid #dee2e6;
            display: grid;
            grid-template-columns: 1fr 1fr auto;
            gap: 15px;
            align-items: end;
        }

        .keyword-stats {
            background: #e3f2fd;
            border: 1px solid #bbdefb;
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .regex-indicator {
            font-size: 0.8rem;
            background: #ffeaa7;
            color: #2d3436;
            padding: 2px 6px;
            border-radius: 4px;
            font-weight: 600;
        }

        .empty-state {
            text-align: center;
            color: #6c757d;
            padding: 40px 20px;
        }

        @media (max-width: 768px) {
            .container {
                padding: 10px;
            }

            .header h1 {
                font-size: 2rem;
            }

            .tab-content {
                padding: 20px;
            }

            .settings-grid {
                grid-template-columns: 1fr;
            }

            .upload-area {
                padding: 40px 15px;
            }

            .status-bar {
                margin: -20px -20px 20px -20px;
            }

            .keyword-item {
                grid-template-columns: 1fr;
                gap: 10px;
            }

            .new-keyword-form {
                grid-template-columns: 1fr;
                gap: 15px;
            }

            .tabs {
                flex-direction: column;
            }
        
        /* Preview Modal Styles */
        #previewModal {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            z-index: 999999;
        }
        
        #previewModal.show {
            display: block;
        }
        
        .modal-overlay {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.9);
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
            box-sizing: border-box;
        }
        
        .modal {
            background: white;
            border-radius: 12px;
            width: 95%;
            max-width: 1600px;
            height: 90vh;
            max-height: 90vh;
            display: flex;
            flex-direction: column;
            box-shadow: 0 20px 60px rgba(0,0,0,0.5);
            position: relative;
        }
        
        .diff-container {
            flex: 1;
            overflow: hidden;
            display: flex;
            background: #e0e0e0;
            gap: 2px;
            min-height: 0;
        }
        
        .document-panel {
            flex: 1;
            display: flex;
            flex-direction: column;
            background: white;
            min-height: 0;
        }
        
        .document-header {
            padding: 15px 20px;
            background: #f8f9fa;
            border-bottom: 2px solid #e0e0e0;
            font-weight: 600;
            flex-shrink: 0;
        }
        
        .document-content {
            flex: 1;
            overflow-y: auto;
            padding: 40px;
            background: white;
            min-height: 0;
        }
        
        .document-page {
            max-width: 650px;
            margin: 0 auto;
            background: white;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            padding: 60px 70px;
            font-family: 'Calibri', sans-serif;
            font-size: 11pt;
            line-height: 1.6;
        }
        
        .paragraph {
            margin-bottom: 12px;
            position: relative;
        }
        
        .paragraph.has-change {
            padding: 8px;
            margin: -8px;
            border-radius: 4px;
        }
        
        .paragraph.has-change::before {
            content: '';
            position: absolute;
            left: -20px;
            top: 0;
            bottom: 0;
            width: 4px;
            background: #667eea;
            border-radius: 2px;
        }
        
        .change-deleted {
            background: #ffe6e6;
            text-decoration: line-through;
            color: #c62828;
            padding: 2px 4px;
            border-radius: 2px;
        }
        
        .change-inserted {
            background: #e6f7e6;
            color: #2e7d32;
            font-weight: 600;
            padding: 2px 4px;
            border-radius: 2px;
        }
        
        .image-block {
            margin: 20px 0;
            padding: 40px;
            background: #f8f9fa;
            border: 2px dashed #ccc;
            border-radius: 8px;
            text-align: center;
            color: #666;
        }
        
        .image-block.removed {
            background: #ffebee;
            border-color: #ef5350;
            color: #c62828;
        }
        
        .legend {
            display: flex;
            gap: 20px;
            font-size: 0.9rem;
        }
        
        .legend-item {
            display: flex;
            align-items: center;
            gap: 6px;
        }
        
        .legend-box {
            width: 20px;
            height: 20px;
            border-radius: 3px;
            border: 1px solid;
        }
        
        .legend-delete {
            background: #ffe6e6;
            border-color: #ffcccc;
        }
        
        .legend-insert {
            background: #e6f7e6;
            border-color: #b3e6b3;
        }
        
        .modal-header {
            padding: 20px 30px;
            border-bottom: 2px solid #e0e0e0;
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border-radius: 12px 12px 0 0;
            flex-shrink: 0;
        }
        
        .close-btn {
            background: none;
            border: none;
            color: white;
            font-size: 2rem;
            cursor: pointer;
            line-height: 1;
            padding: 0;
            width: 32px;
            height: 32px;
            border-radius: 50%;
            transition: background 0.2s;
        }
        
        .close-btn:hover {
            background: rgba(255,255,255,0.2);
        }
        
        .statistics-bar {
            padding: 15px 30px;
            background: #f8f9fa;
            border-bottom: 1px solid #e0e0e0;
            display: flex;
            justify-content: space-around;
            flex-wrap: wrap;
            gap: 20px;
            flex-shrink: 0;
        }
        
        .stat-item {
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 5px;
        }
        
        .stat-value {
            font-size: 1.5rem;
            font-weight: 700;
            color: #667eea;
        }
        
        .stat-label {
            font-size: 0.9rem;
            color: #666;
        }
        
        .modal-footer {
            padding: 20px 30px;
            border-top: 2px solid #e0e0e0;
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: #f8f9fa;
            border-radius: 0 0 12px 12px;
            flex-shrink: 0;
        }
        
        .action-buttons {
            display: flex;
            gap: 15px;
        }
    </style>
</head>
<body>
    <div class="container">
        <header class="header">
            <h1>File Blinder</h1>
            <p>Secure Document Anonymization Tool - Remove images, document coloring and replace sensitive keywords</p>
        </header>

        <div class="main-content">
            <div class="status-bar">
                <span>Local Web Server</span>
                <span class="version">v1.1</span>
            </div>

            <div class="tabs">
                <button class="tab active" data-tab="upload">Upload & Process</button>
                <button class="tab" data-tab="keywords">Keyword Management</button>
                <button class="tab" data-tab="settings">Settings</button>
            </div>

            <!-- Upload Tab -->
            <div class="tab-content active" id="upload">
                <form id="uploadForm" enctype="multipart/form-data">
                    <div class="upload-area" id="uploadArea">
                        <div class="upload-icon">📁</div>
                        <div class="upload-text">Drag and drop your file here</div>
                        <div class="upload-subtext">or click to select a file</div>
                        <input type="file" id="fileInput" name="file" class="file-input" accept=".docx,.html,.htm,.txt">
                    </div>

                    <div class="supported-formats">
                        <div class="format-badge">📄 DOCX</div>
                        <div class="format-badge">📝 TXT</div>
                    </div>

                    <div class="file-info" id="fileInfo">
                        <h3>📄 File Information</h3>
                        <div id="fileDetails"></div>
                    </div>

                    <div class="processing-status" id="processingStatus">
                        🔄 Processing your file...
                    </div>

                    <div class="progress" id="progressBar">
                        <div class="progress-bar" id="progressFill"></div>
                    </div>

                    <div class="results" id="results"></div>

                    <div style="text-align: center; margin-top: 30px;">
                        <button type="button" class="btn btn-secondary" id="previewBtn" disabled style="margin-right: 10px;" onclick="showPreview()">
                            Preview Changes
                        </button>
                        <button type="submit" class="btn" id="processBtn" disabled>Blind File</button>
                    </div>
                </form>
            </div>
            
            <!-- Keywords Tab -->
            <div class="tab-content" id="keywords">
                <h2>🔤 Keyword Management</h2>
                <p style="color: #6c757d; margin-bottom: 20px;">Manage words and patterns to be replaced during file processing. Changes are automatically saved.</p>

                <div class="keyword-stats" id="keywordStats">
                    <span>Total Keywords: <strong id="totalKeywords">0</strong></span>
                    <span>Active: <strong id="activeKeywords">0</strong></span>
                    <button class="btn btn-secondary btn-small" onclick="resetToDefaults()">Reset to Defaults</button>
                </div>

                <div class="keywords-container">
                    <div class="keyword-header">
                        <span>Manage Keywords</span>
                        <button class="btn btn-success btn-small" onclick="exportKeywords()">Export Keywords</button>
                    </div>

                    <div class="keyword-list" id="keywordList">
                        <!-- Keywords will be loaded here -->
                    </div>

                    <div class="new-keyword-form">
                        <div class="form-group">
                            <input type="text" id="newKeywordOriginal" class="keyword-input" placeholder="Word or pattern to replace" />
                        </div>
                        <div class="form-group">
                            <input type="text" id="newKeywordReplacement" class="keyword-input" placeholder="Replacement text" value="REDACTED" />
                        </div>
                        <button type="button" class="btn btn-success btn-small" onclick="addKeyword()">Add Keyword</button>
                    </div>
                </div>
            </div>

            <!-- Settings Tab -->
            <div class="tab-content" id="settings">
                <h2>⚙️ Processing Settings</h2>

                <div class="settings-section">
                    <h3>📄 Document Processing</h3>

                    <div class="settings-grid">
                        <div>
                            <div class="form-group">
                                <label for="processingMethod">Processing Method (DOCX only)</label>
                                <select id="processingMethod" name="processing_method">
                                    <option value="safe">Safe (Recommended)</option>
                                    <option value="xml">XML Direct Processing</option>
                                </select>
                            </div>

                            <div class="checkbox-group">
                                <input type="checkbox" id="standardizeFormatting" name="standardize_formatting" checked>
                                <label for="standardizeFormatting">Standardize formatting</label>
                            </div>

                            <div class="checkbox-group">
                                <input type="checkbox" id="fontColorBlack" name="font_color_black" checked>
                                <label for="fontColorBlack">Force all text to black</label>
                            </div>
                        </div>

                        <div>
                            <div class="form-group">
                                <label for="fontName">Font Name</label>
                                <input type="text" id="fontName" name="font_name" value="Calibri">
                            </div>

                            <div class="form-group">
                                <label for="fontSize">Font Size</label>
                                <input type="number" id="fontSize" name="font_size" value="11" min="8" max="72">
                            </div>

                            <div class="checkbox-group">
                                <input type="checkbox" id="removeShading" name="remove_shading" checked>
                                <label for="removeShading">Remove background shading</label>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    
    <footer style="text-align: center; padding: 30px 20px; margin-top: 20px;">
        <img src="/logo.svg" alt="Company Logo" style="max-width: 150px; max-height: 60px; height: auto; opacity: 0.8;">
    </footer>

    <!-- Preview Modal - MUST BE AT END OF BODY, OUTSIDE ALL OTHER CONTAINERS -->
    <div id="previewModal">
        <div class="modal-overlay">
            <div class="modal">
                <div class="modal-header">
                    <h2 id="previewTitle">Preview Changes</h2>
                    <button class="close-btn" onclick="closePreview()">&times;</button>
                </div>
    
                <div class="statistics-bar" id="previewStats">
                    <!-- Statistics will be inserted here -->
                </div>
    
                <div class="diff-container">
                    <div class="document-panel">
                        <div class="document-header">
                            <span>Original Document</span>
                        </div>
                        <div class="document-content" id="originalDocContent">
                            <div style="padding: 40px; text-align: center; color: #666;">
                                Loading preview...
                            </div>
                        </div>
                    </div>
    
                    <div class="document-panel">
                        <div class="document-header">
                            <span>After Processing (Blinded)</span>
                        </div>
                        <div class="document-content" id="processedDocContent">
                            <div style="padding: 40px; text-align: center; color: #666;">
                                Loading preview...
                            </div>
                        </div>
                    </div>
                </div>
    
                <div class="modal-footer">
                    <div class="legend">
                        <div class="legend-item">
                            <div class="legend-box legend-delete"></div>
                            <span>Removed</span>
                        </div>
                        <div class="legend-item">
                            <div class="legend-box legend-insert"></div>
                            <span>Inserted</span>
                        </div>
                    </div>
                    <div class="action-buttons">
                        <button class="btn btn-secondary" onclick="closePreview()">Cancel</button>
                        <button class="btn" onclick="proceedWithProcessing()">Looks Good - Process File</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <script>
        let selectedFile = null;
        let keywords = [];
        let previewData = null;

        document.addEventListener('DOMContentLoaded', function() {
            initializeTabs();
            initializeUpload();
            setupForm();
            loadKeywords();
            const previewModal = document.getElementById('previewModal');
    if (previewModal) {
        previewModal.style.display = 'none';
    }
        });

        function initializeTabs() {
            const tabs = document.querySelectorAll('.tab');
            const tabContents = document.querySelectorAll('.tab-content');

            tabs.forEach(tab => {
                tab.addEventListener('click', () => {
                    tabs.forEach(t => t.classList.remove('active'));
                    tabContents.forEach(tc => tc.classList.remove('active'));

                    tab.classList.add('active');
                    const targetId = tab.getAttribute('data-tab');
                    document.getElementById(targetId).classList.add('active');
                });
            });
        }

        function initializeUpload() {
            const uploadArea = document.getElementById('uploadArea');
            const fileInput = document.getElementById('fileInput');

            uploadArea.addEventListener('click', () => fileInput.click());
            fileInput.addEventListener('change', handleFileSelect);

            uploadArea.addEventListener('dragover', (e) => {
                e.preventDefault();
                uploadArea.classList.add('drag-over');
            });

            uploadArea.addEventListener('dragleave', (e) => {
                if (!uploadArea.contains(e.relatedTarget)) {
                    uploadArea.classList.remove('drag-over');
                }
            });

            uploadArea.addEventListener('drop', (e) => {
                e.preventDefault();
                uploadArea.classList.remove('drag-over');
                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    fileInput.files = files;
                    handleFileSelect({ target: { files: files } });
                }
            });
        }

        function handleFileSelect(event) {
            const file = event.target.files[0];
            if (!file) return;

            const supportedTypes = ['.docx', '.html', '.htm', '.txt'];
            const fileExtension = '.' + file.name.split('.').pop().toLowerCase();

            if (!supportedTypes.includes(fileExtension)) {
                showResults(`Unsupported file type: ${fileExtension}<br>Supported formats: DOCX, HTML, TXT`, 'error');
                return;
            }

            selectedFile = file;
            displayFileInfo(file);
            document.getElementById('processBtn').disabled = false;
            document.getElementById('previewBtn').disabled = false;
            hideResults();
        }

        function displayFileInfo(file) {
            const fileInfo = document.getElementById('fileInfo');
            const fileDetails = document.getElementById('fileDetails');

            const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
            const fileSize = formatFileSize(file.size);
            const fileType = getFileTypeDescription(fileExtension);

            fileDetails.innerHTML = `
                <div class="info-grid">
                    <div class="info-item"><strong>Name:</strong><br>${file.name}</div>
                    <div class="info-item"><strong>Type:</strong><br>${fileType}</div>
                    <div class="info-item"><strong>Size:</strong><br>${fileSize}</div>
                    <div class="info-item"><strong>Last Modified:</strong><br>${new Date(file.lastModified).toLocaleDateString()}</div>
                </div>
            `;

            fileInfo.classList.add('show');
        }

        function getFileTypeDescription(extension) {
            switch(extension) {
                case '.docx': return 'Microsoft Word Document';
                case '.html':
                case '.htm': return 'HTML Document';
                case '.txt': return 'Plain Text File';
                default: return 'Unknown';
            }
        }

        function formatFileSize(bytes) {
            if (bytes === 0) return '0 Bytes';
            const k = 1024;
            const sizes = ['Bytes', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        }

        function setupForm() {
            const form = document.getElementById('uploadForm');
            form.addEventListener('submit', async (e) => {
                e.preventDefault();

                if (!selectedFile) {
                    showResults('Please select a file first.', 'error');
                    return;
                }

                const formData = new FormData(form);

                showProcessingStatus();
                updateProgress(10);

                try {
                    const response = await fetch('/process', {
                        method: 'POST',
                        body: formData
                    });

                    updateProgress(60);

                    if (!response.ok) {
                        const errorData = await response.json();
                        throw new Error(errorData.error || 'Processing failed');
                    }

                    updateProgress(90);

                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;

                    const contentDisposition = response.headers.get('Content-Disposition');
                    const filename = contentDisposition 
                        ? contentDisposition.split('filename=')[1].replace(/"/g, '')
                        : selectedFile.name.replace(/\.[^/.]+$/, '_blinded$&');

                    a.download = filename;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);

                    updateProgress(100);
                    showResults(`File processed successfully!<br>Download started: <strong>${filename}</strong>`, 'success');

                } catch (error) {
                    showResults(`Error processing file:<br>${error.message}`, 'error');
                } finally {
                    hideProcessingStatus();
                    setTimeout(() => updateProgress(0), 2000);
                }
            });
        }

        // Keyword Management Functions
        async function loadKeywords() {
            try {
                const response = await fetch('/keywords');
                keywords = await response.json();
                renderKeywords();
                updateKeywordStats();
            } catch (error) {
                console.error('Error loading keywords:', error);
            }
        }

        function renderKeywords() {
            const keywordList = document.getElementById('keywordList');

            if (keywords.length === 0) {
                keywordList.innerHTML = '<div class="empty-state">No keywords configured. Add keywords below to start redacting content.</div>';
                return;
            }

            keywordList.innerHTML = keywords.map((keyword, index) => `
                <div class="keyword-item">
                    <input type="checkbox" class="keyword-toggle" 
                           ${keyword.enabled ? 'checked' : ''} 
                           onchange="toggleKeyword(${index})">
                    <input type="text" class="keyword-input" 
                           value="${escapeHtml(keyword.original)}" 
                           onchange="updateKeyword(${index}, 'original', this.value)"
                           placeholder="Original text">
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <input type="text" class="keyword-input" 
                               value="${escapeHtml(keyword.replacement)}" 
                               onchange="updateKeyword(${index}, 'replacement', this.value)"
                               placeholder="Replacement">
                        ${keyword.original.startsWith('\\b') || keyword.original.includes('[') ? 
                          '<span class="regex-indicator">REGEX</span>' : ''}
                    </div>
                    <div class="keyword-actions">
                        <button class="btn btn-danger btn-small" onclick="removeKeyword(${index})">Remove</button>
                    </div>
                </div>
            `).join('');
        }

        function updateKeywordStats() {
            const total = keywords.length;
            const active = keywords.filter(k => k.enabled).length;

            document.getElementById('totalKeywords').textContent = total;
            document.getElementById('activeKeywords').textContent = active;
        }

        async function addKeyword() {
            const original = document.getElementById('newKeywordOriginal').value.trim();
            const replacement = document.getElementById('newKeywordReplacement').value.trim();

            if (!original || !replacement) {
                alert('Please enter both original text and replacement.');
                return;
            }

            const newKeyword = {
                original: original,
                replacement: replacement,
                enabled: true
            };

            keywords.push(newKeyword);
            await saveKeywords();
            renderKeywords();
            updateKeywordStats();

            document.getElementById('newKeywordOriginal').value = '';
            document.getElementById('newKeywordReplacement').value = 'REDACTED';
        }

        async function removeKeyword(index) {
            if (confirm('Are you sure you want to remove this keyword?')) {
                keywords.splice(index, 1);
                await saveKeywords();
                renderKeywords();
                updateKeywordStats();
            }
        }

        async function toggleKeyword(index) {
            keywords[index].enabled = !keywords[index].enabled;
            await saveKeywords();
            updateKeywordStats();
        }

        async function updateKeyword(index, field, value) {
            keywords[index][field] = value;
            await saveKeywords();
            renderKeywords();
        }

        async function saveKeywords() {
            try {
                const response = await fetch('/keywords', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(keywords)
                });

                if (!response.ok) {
                    throw new Error('Failed to save keywords');
                }
            } catch (error) {
                console.error('Error saving keywords:', error);
                alert('Error saving keywords. Changes may not be preserved.');
            }
        }

        async function resetToDefaults() {
            if (confirm('This will reset all keywords to default values. Are you sure?')) {
                try {
                    const response = await fetch('/keywords/reset', { method: 'POST' });
                    if (response.ok) {
                        await loadKeywords();
                        alert('Keywords reset to defaults successfully.');
                    } else {
                        throw new Error('Failed to reset keywords');
                    }
                } catch (error) {
                    console.error('Error resetting keywords:', error);
                    alert('Error resetting keywords.');
                }
            }
        }

        function exportKeywords() {
            const dataStr = JSON.stringify(keywords, null, 2);
            const dataBlob = new Blob([dataStr], {type: 'application/json'});
            const url = URL.createObjectURL(dataBlob);
            const link = document.createElement('a');
            link.href = url;
            link.download = 'file-blinder-keywords.json';
            link.click();
            URL.revokeObjectURL(url);
        }

        function escapeHtml(text) {
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }

        function showProcessingStatus() {
            document.getElementById('processingStatus').classList.add('show');
            document.getElementById('progressBar').classList.add('show');
            document.getElementById('processBtn').disabled = true;
        }

        function hideProcessingStatus() {
            document.getElementById('processingStatus').classList.remove('show');
            document.getElementById('processBtn').disabled = false;
        }

        function updateProgress(percent) {
            document.getElementById('progressFill').style.width = percent + '%';
        }

        function showResults(message, type = 'success') {
            const results = document.getElementById('results');
            results.innerHTML = message;
            results.className = `results show ${type}`;
        }

        function hideResults() {
            document.getElementById('results').classList.remove('show');
        }

        // PREVIEW FUNCTIONS
        // PREVIEW FUNCTIONS
        async function showPreview() {
            if (!selectedFile) {
                showResults('Please select a file first.', 'error');
                return;
            }
        
            const formData = new FormData(document.getElementById('uploadForm'));
            
            showProcessingStatus();
            updateProgress(10);
        
            try {
                updateProgress(30);
                const response = await fetch('/preview', {
                    method: 'POST',
                    body: formData
                });
        
                updateProgress(70);
        
                if (!response.ok) {
                    const errorData = await response.json();
                    throw new Error(errorData.error || 'Preview generation failed');
                }
        
                previewData = await response.json();
                console.log('Preview data received:', previewData);
                
                updateProgress(90);
                
                // Render preview BEFORE showing modal
                renderPreview(previewData);
                
                updateProgress(95);
                
                // Small delay to ensure rendering is complete
                await new Promise(resolve => setTimeout(resolve, 100));
                
                // Now show the modal - use BOTH display and class
                const modal = document.getElementById('previewModal');
                modal.style.display = 'block';  // Make it part of layout
                setTimeout(() => {
                    modal.classList.add('show');  // Then fade it in
                }, 10);
                document.body.style.overflow = 'hidden';
                
                console.log('Modal should now be visible');
                
                updateProgress(100);
                
            } catch (error) {
                console.error('Preview error:', error);
                showResults(`Error generating preview: ${error.message}`, 'error');
            } finally {
                hideProcessingStatus();
                setTimeout(() => updateProgress(0), 1000);
            }
        }
        
        function closePreview() {
            console.log('Closing preview modal');
            const modal = document.getElementById('previewModal');
            modal.classList.remove('show');
            // Wait for fade out animation before hiding completely
            setTimeout(() => {
                modal.style.display = 'none';
            }, 200);
            document.body.style.overflow = 'auto';
        }

         function renderPreview(data) {
            console.log('Rendering preview with data:', data);
            
            document.getElementById('previewTitle').textContent = `Preview Changes - ${data.filename}`;
            
            const statsHtml = `
                <div class="stat-item">
                    <span class="stat-value">${data.statistics.images_removed}</span>
                    <span class="stat-label">Images Removed</span>
                </div>
                <div class="stat-item">
                    <span class="stat-value">${data.statistics.keywords_replaced}</span>
                    <span class="stat-label">Keywords Replaced</span>
                </div>
                <div class="stat-item">
                    <span class="stat-value">${data.statistics.formatting_changes}</span>
                    <span class="stat-label">Formatting Standardized</span>
                </div>
                <div class="stat-item">
                    <span class="stat-value">${data.statistics.table_shadings_removed}</span>
                    <span class="stat-label">Table Shadings Removed</span>
                </div>
            `;
            document.getElementById('previewStats').innerHTML = statsHtml;
            
            console.log('Rendering original document...');
            renderDocument(data.original, data.diff, 'original', document.getElementById('originalDocContent'));
            
            console.log('Rendering processed document...');
            renderDocument(data.processed, data.diff, 'processed', document.getElementById('processedDocContent'));
            
            console.log('Setting up scroll sync...');
            setupSyncScroll();
            
            console.log('Preview rendering complete');
        }

        function renderDocument(structure, diff, side, container) {
            const page = document.createElement('div');
            page.className = 'document-page';
            
            structure.paragraphs.forEach((para, idx) => {
                const paraDiv = document.createElement('div');
                paraDiv.className = 'paragraph';
                
                const paraChange = diff.paragraph_changes.find(pc => pc.index === idx);
                const imageChange = diff.image_changes.find(ic => ic.paragraph_index === idx);
                
                if (paraChange || imageChange) {
                    paraDiv.classList.add('has-change');
                }
                
                if (para.has_image) {
                    const imageDiv = document.createElement('div');
                    imageDiv.className = side === 'original' ? 'image-block' : 'image-block removed';
                    imageDiv.innerHTML = side === 'original' 
                        ? '<strong>Image in Document</strong>'
                        : '<strong>Image Removed</strong><div style="margin-top: 8px;">This image will not appear in the processed document</div>';
                    page.appendChild(imageDiv);
                }
                
                let textHtml = escapeHtml(para.text);
                
                if (paraChange) {
                    paraChange.changes.forEach(change => {
                        if (change.type === 'replace') {
                            const original = escapeHtml(change.original);
                            const processed = escapeHtml(change.processed);
                            
                            if (side === 'original') {
                                textHtml = textHtml.replace(
                                    original,
                                    `<span class="change-deleted">${original}</span>`
                                );
                            } else {
                                textHtml = textHtml.replace(
                                    processed,
                                    `<span class="change-inserted">${processed}</span>`
                                );
                            }
                        }
                    });
                }
                
                if (para.style && para.style.includes('Heading')) {
                    const level = para.style.match(/\d+/);
                    if (level) {
                        const tag = `h${level[0]}`;
                        const heading = document.createElement(tag);
                        heading.innerHTML = textHtml;
                        page.appendChild(heading);
                        return;
                    }
                }
                
                paraDiv.innerHTML = textHtml;
                page.appendChild(paraDiv);
            });
            
            container.innerHTML = '';
            container.appendChild(page);
        }

        function setupSyncScroll() {
            const original = document.getElementById('originalDocContent');
            const processed = document.getElementById('processedDocContent');
            let isScrolling = false;

            original.addEventListener('scroll', () => {
                if (!isScrolling) {
                    isScrolling = true;
                    processed.scrollTop = original.scrollTop;
                    setTimeout(() => isScrolling = false, 50);
                }
            });

            processed.addEventListener('scroll', () => {
                if (!isScrolling) {
                    isScrolling = true;
                    original.scrollTop = processed.scrollTop;
                    setTimeout(() => isScrolling = false, 50);
                }
            });
        }

        function proceedWithProcessing() {
            closePreview();
            document.getElementById('uploadForm').dispatchEvent(new Event('submit'));
        }
    </script>
    
</body>
</html>
"""


@app.route('/')
def index():
    """Serve the main page"""
    return render_template_string(HTML_TEMPLATE)


@app.route('/keywords', methods=['GET', 'POST'])
def manage_keywords():
    """Manage keywords - GET to retrieve, POST to save"""
    if request.method == 'GET':
        keywords = load_keywords()
        return jsonify(keywords)

    elif request.method == 'POST':
        try:
            new_keywords = request.get_json()
            if not isinstance(new_keywords, list):
                return jsonify({'error': 'Invalid keyword format'}), 400

            # Validate keyword structure
            for keyword in new_keywords:
                if not isinstance(keyword, dict) or 'original' not in keyword or 'replacement' not in keyword:
                    return jsonify({'error': 'Invalid keyword structure'}), 400

            success = save_keywords(new_keywords)
            if success:
                return jsonify({'success': True})
            else:
                return jsonify({'error': 'Failed to save keywords'}), 500

        except Exception as e:
            return jsonify({'error': str(e)}), 500


@app.route('/keywords/reset', methods=['POST'])
def reset_keywords():
    """Reset keywords to defaults"""
    try:
        # Delete existing keywords file to force reload of defaults
        if KEYWORDS_FILE.exists():
            KEYWORDS_FILE.unlink()

        # Load defaults (which will be saved automatically)
        default_keywords = load_keywords()
        return jsonify({'success': True, 'keywords': default_keywords})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/process', methods=['POST'])
def process_file():
    """Process the uploaded file"""
    if not FILEBLINDER_AVAILABLE:
        return jsonify(
            {'error': 'FileBlinder module not available. Please ensure file_blinder.py is in the same directory.'}), 500

    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        if not allowed_file(file.filename):
            return jsonify(
                {'error': f'File type not supported. Supported formats: {", ".join(ALLOWED_EXTENSIONS)}'}), 400

        # Secure the filename
        filename = secure_filename(file.filename)

        # Get form parameters
        processing_method = request.form.get('processing_method', 'safe')
        standardize_formatting = request.form.get('standardize_formatting') == 'on'
        font_name = request.form.get('font_name', 'Calibri')
        try:
            font_size = int(request.form.get('font_size', 11))
        except ValueError:
            font_size = 11
        font_color_black = request.form.get('font_color_black') == 'on'
        remove_shading = request.form.get('remove_shading') == 'on'

        # Create temporary files
        with tempfile.NamedTemporaryFile(delete=False, suffix=Path(filename).suffix) as tmp_input:
            file.save(tmp_input.name)
            input_path = tmp_input.name

        # Generate output filename
        original_name = Path(filename).stem
        original_ext = Path(filename).suffix
        output_filename = f"{original_name}_blinded{original_ext}"

        with tempfile.NamedTemporaryFile(delete=False, suffix=original_ext) as tmp_output:
            output_path = tmp_output.name

        try:
            # Get active keywords from persistent storage
            active_keywords = get_active_keywords()

            # Initialize FileBlinder with current keywords
            blinder = FileBlinder(
                keyword_replacements=active_keywords,
                standardize_formatting=standardize_formatting,
                font_name=font_name,
                font_size=font_size,
                font_color_black=font_color_black,
                grey_shading=False  # Always remove shading for security
            )

            # Process the file
            result = blinder.blind_file(
                input_path=input_path,
                output_path=output_path,
                method=processing_method
            )

            if not result:
                return jsonify({'error': 'File processing failed - check console for details'}), 500

            # Send the processed file
            return send_file(
                output_path,
                as_attachment=True,
                download_name=output_filename,
                mimetype='application/octet-stream'
            )

        finally:
            # Clean up input file immediately
            try:
                os.unlink(input_path)
            except:
                pass

    except Exception as e:
        # Log the error for debugging
        print(f"Error processing file: {e}")
        print(traceback.format_exc())
        return jsonify({'error': f'Processing error: {str(e)}'}), 500


@app.route('/status')
def status():
    """Simple status endpoint"""
    keywords = load_keywords()
    return jsonify({
        'status': 'running',
        'fileblinder_available': FILEBLINDER_AVAILABLE,
        'supported_formats': list(ALLOWED_EXTENSIONS),
        'keywords_count': len(keywords),
        'active_keywords': len([k for k in keywords if k.get('enabled', True)])
    })

@app.route('/logo.svg')
def serve_logo():
    """Serve the company logo"""
    logo_path = current_dir / 'logo.svg'
    if logo_path.exists():
        return send_file(logo_path, mimetype='image/svg+xml')
    else:
        # Return a placeholder if logo not found
        return '', 404
def open_browser():
    """Open the default web browser after a short delay"""
    import time
    time.sleep(1.5)  # Wait for server to start
    webbrowser.open('http://localhost:5000')


def main():
    """Main function to run the server"""
    print("=" * 60)
    print("🔒 FILE BLINDER - Local Web Server v1.1")
    print("=" * 60)
    print()

    if not FILEBLINDER_AVAILABLE:
        print("❌ ERROR: file_blinder.py not found!")
        print("Please ensure file_blinder.py is in the same directory as this script.")
        print()
        input("Press Enter to exit...")
        return

    # Initialize keywords storage
    keywords = load_keywords()
    print(f"FileBlinder module loaded successfully")
    print(f"Keywords loaded: {len(keywords)} total ({len([k for k in keywords if k.get('enabled', True)])} active)")
    print(f"Keywords storage: {KEYWORDS_FILE}")
    print("Starting local web server...")
    print()
    print("🔗 Open your browser and go to: http://localhost:5000")
    print("🛑 Press Ctrl+C to stop the server")
    print()
    print("Features available:")
    print("  📁 Drag & drop file upload")
    print("  🔤 Keyword management with persistent storage")
    print("  🔧 Customizable processing settings")
    print("  📄 Support for DOCX, HTML, and TXT files")
    print("  🔒 Complete anonymization and formatting")
    print("  📥 Automatic file download")
    print()

    # Open browser in a separate thread
    browser_thread = threading.Thread(target=open_browser)
    browser_thread.daemon = True
    browser_thread.start()

    try:
        # Run the Flask app
        app.run(
            debug=False,  # Set to True for development
            host='127.0.0.1',  # Only accessible locally for security
            port=5000,
            use_reloader=False  # Disable reloader to prevent double browser opening
        )
    except KeyboardInterrupt:
        print("\nServer stopped by user")
    except Exception as e:
        print(f"\nServer error: {e}")
    finally:
        print("Thank you for using File Blinder!")


if __name__ == '__main__':
    main()