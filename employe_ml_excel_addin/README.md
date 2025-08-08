# Employee Churn Predictor Excel Add-in

This Excel add-in uses machine learning to predict employee churn probability. It connects to a Flask server running a trained XGBoost model.

## Features

- Load sample employee data directly into Excel
- Select data rows and get churn predictions
- Display predictions in a new column with proper formatting
- User-friendly interface with clear instructions

## Prerequisites

1. **Node.js** (version 14 or higher)
2. **Python** (version 3.7 or higher) with Flask
3. **Excel** (desktop version)
4. **Office Add-in development tools**

## Setup Instructions

### 1. Install Office Add-in Development Tools

```bash
npm install -g office-addin-dev-certs
npm install -g office-addin-debugging
```

### 2. Install Dependencies

```bash
cd Excel_web_app/employe_ml_excel_addin
npm install
```

### 3. Start the Flask Server

```bash
cd ../flask_server
python app.py
```

The Flask server will run on `http://localhost:5000`

### 4. Start the Excel Add-in

```bash
cd ../employe_ml_excel_addin
npm start
```

This will:
- Start the development server on `https://localhost:3000`
- Open Excel with the add-in loaded
- Trust the development certificate

## How to Use

1. **Load Sample Data**: Click the "Load Sample Data" button to populate Excel with example employee data
2. **Select Data**: Select the data rows (excluding the header row)
3. **Predict**: Click "Predict Churn" to get predictions
4. **View Results**: Predictions will appear in a new column with churn probabilities

## Data Format

The add-in expects data in the following format:
- `satisfaction_level` (0-1)
- `last_evaluation` (0-1)
- `number_project` (integer)
- `average_montly_hours` (integer)
- `time_spend_company` (integer)
- `Work_accident` (0 or 1)
- `promotion_last_5years` (0 or 1)
- `Department` (string: sales, technical, support, IT, etc.)
- `salary` (string: low, medium, high)

## Troubleshooting

### Certificate Issues
If you encounter certificate errors:
```bash
office-addin-dev-certs install
```

### Port Conflicts
If ports 3000 or 5000 are in use, modify the configuration in:
- `package.json` for the webpack dev server port
- `app.py` for the Flask server port

### CORS Issues
The Flask server includes CORS headers, but if you encounter issues, ensure the server is running on `http://localhost:5000`

## Development

- The main component is in `src/taskpane/components/App.jsx`
- The Flask server is in `../flask_server/app.py`
- The trained model is `../flask_server/XGBoost.pkl`

## Building for Production

```bash
npm run build
```

This creates a production build in the `dist` folder that can be deployed to a web server.
