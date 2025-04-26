from flask import Flask, request, jsonify
from scanner.nmap_scanner import run_nmap_scan
from ml_model.vulnerability_predictor import predict_vulnerabilities

app = Flask(__name__)

@app.route('/scan', methods=['POST'])
def scan():
    data = request.get_json()
    target_ip = data.get('target_ip')
    if not target_ip:
        return jsonify({'error': 'Missing target_ip'}), 400

    scan_data = run_nmap_scan(target_ip)
    predictions = predict_vulnerabilities(scan_data)

    return jsonify({
        'target': target_ip,
        'scan_data': scan_data,
        'vulnerability_predictions': predictions
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)