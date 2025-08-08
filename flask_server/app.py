from flask import Flask, request, jsonify
from flask_cors import CORS
import numpy as np
import joblib


app = Flask(__name__)
CORS(app)

# Load trained model
model = joblib.load("XGBoost.pkl")

# Encoding dictionaries
department_encode = {
    'sales': 7,
    'technical': 9,
    'support': 8,
    'IT': 0,
    'product_mng': 6,
    'marketing': 5,
    'RandD': 1,
    'accounting': 2,
    'hr': 3,
    'management': 4
}

salary_encode = {
    'low': 1,
    'medium': 2,
    'high': 3
}



@app.route("/predict", methods=["POST"])
def predict():
    try:
        json_data = request.json["data"]  # expect: list of dicts
        feature_list = []

        for record in json_data:
            features = [
                record['satisfaction_level'],
                record['last_evaluation'],
                record['number_project'],
                record['average_montly_hours'],
                record['time_spend_company'],
                record['Work_accident'],
                record['promotion_last_5years'],
                department_encode.get(record['Department'], -1),
                salary_encode.get(record['salary'], -1)
            ]
            feature_list.append(features)

        array = np.array(feature_list)
        predictions = model.predict_proba(array)

        return jsonify({"predictions": predictions.tolist()})
    
    except Exception as e:
        return jsonify({"error": str(e)}), 400

if __name__ == '__main__':
    app.run(debug=True,port=5000)