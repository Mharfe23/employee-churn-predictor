import requests
import json

def test_flask_server():
    """Test the Flask server with sample data"""
    
    # Sample data matching the expected format
    test_data = {
        "data": [
            {
                "satisfaction_level": 0.38,
                "last_evaluation": 0.53,
                "number_project": 2,
                "average_montly_hours": 157,
                "time_spend_company": 3,
                "Work_accident": 0,
                "promotion_last_5years": 0,
                "Department": "sales",
                "salary": "low"
            },
            {
                "satisfaction_level": 0.8,
                "last_evaluation": 0.86,
                "number_project": 5,
                "average_montly_hours": 262,
                "time_spend_company": 6,
                "Work_accident": 0,
                "promotion_last_5years": 0,
                "Department": "technical",
                "salary": "medium"
            }
        ]
    }
    
    try:
        # Send POST request to the Flask server
        response = requests.post(
            "http://localhost:5000/predict",
            headers={"Content-Type": "application/json"},
            data=json.dumps(test_data)
        )
        
        if response.status_code == 200:
            result = response.json()
            print("✅ Flask server is working correctly!")
            print(f"Response: {result}")
            
            # Display predictions in a readable format
            predictions = result.get("predictions", [])
            for i, pred in enumerate(predictions):
                stay_prob = pred[0]
                churn_prob = pred[1]
                print(f"Employee {i+1}: Stay Probability: {stay_prob:.4f}, Churn Probability: {churn_prob:.4f}")
                
        else:
            print(f"❌ Server returned status code: {response.status_code}")
            print(f"Response: {response.text}")
            
    except requests.exceptions.ConnectionError:
        print("❌ Could not connect to Flask server. Make sure it's running on http://localhost:5000")
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    print("Testing Flask server...")
    test_flask_server()
