import * as React from "react";
import { useState } from "react";

const App = () => {
  const [message, setMessage] = useState("Select data rows and click Predict");
  const [isLoading, setIsLoading] = useState(false);

  const handlePredict = async () => {
    setIsLoading(true);
    setMessage("Processing...");
    
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = context.workbook.getSelectedRange();
        
        // Load the range values
        range.load("values, address");
        await context.sync();
        
        console.log(`Selected range: ${range.address}`);
        console.log(`Number of rows: ${range.values.length}`);
        
        if (range.values.length === 0) {
          setMessage("Please select some data rows first");
          setIsLoading(false);
          return;
        }

        // Map each row to a dictionary for the ML model
        const records = range.values.map((row) => ({
          satisfaction_level: parseFloat(row[0]) || 0,
          last_evaluation: parseFloat(row[1]) || 0,
          number_project: parseInt(row[2]) || 0,
          average_montly_hours: parseInt(row[3]) || 0,
          time_spend_company: parseInt(row[4]) || 0,
          Work_accident: parseInt(row[5]) || 0,
          promotion_last_5years: parseInt(row[6]) || 0,
          Department: String(row[7] || ''),
          salary: String(row[8] || '')
        }));

        console.log("Sending data to server:", records);

        const response = await fetch("http://localhost:5000/predict", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ data: records })
        });

        const result = await response.json();
        
        if (result.error) {
          setMessage(`Error: ${result.error}`);
        } else {
          // Get the predictions
          const predictions = result.predictions;
          
          // Find the next empty column to write predictions
          const usedRange = sheet.getUsedRange();
          usedRange.load("columnCount");
          await context.sync();
          
          const nextColumn = usedRange.columnCount + 1;
          const predictionColumn = String.fromCharCode(64 + nextColumn); // Convert to letter (A, B, C, etc.)
          
          // Create range for predictions (same number of rows as selected data)
          const predictionRange = sheet.getRange(`${predictionColumn}1:${predictionColumn}${range.values.length}`);
          
          // Format predictions as percentages and write to Excel
          const formattedPredictions = predictions.map(pred => {
            const churnProbability = pred[1]; // Probability of leaving (churn)
            return [churnProbability.toFixed(4)]; // Format as 4 decimal places
          });
          
          predictionRange.values = formattedPredictions;
          
          // Add header for the prediction column
          const headerRange = sheet.getRange(`${predictionColumn}1`);
          headerRange.values = [["Churn Probability"]];
          
          // Format the header
          headerRange.format.fill.color = "#4472C4";
          headerRange.format.font.color = "white";
          headerRange.format.font.bold = true;
          
          // Format prediction values as percentages
          predictionRange.format.numberFormat = "0.00%";
          
          setMessage(`Predictions added in column ${predictionColumn}!`);
        }
        
      });
    } catch (err) {
      console.error("Error:", err);
      setMessage("Something went wrong: " + err.message);
    } finally {
      setIsLoading(false);
    }
  };

  const handleLoadSampleData = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        // Sample data headers
        const headers = [
          ["satisfaction_level", "last_evaluation", "number_project", "average_montly_hours", 
           "time_spend_company", "Work_accident", "promotion_last_5years", "Department", "salary"]
        ];
        
        // Sample data rows
        const sampleData = [
          [0.38, 0.53, 2, 157, 3, 0, 0, "sales", "low"],
          [0.8, 0.86, 5, 262, 6, 0, 0, "sales", "medium"],
          [0.11, 0.88, 7, 272, 4, 0, 0, "technical", "medium"],
          [0.72, 0.87, 5, 223, 5, 0, 0, "support", "low"],
          [0.37, 0.52, 2, 159, 3, 0, 0, "IT", "low"]
        ];
        
        // Write headers
        const headerRange = sheet.getRange("A1:I1");
        headerRange.values = headers;
        headerRange.format.fill.color = "#4472C4";
        headerRange.format.font.color = "white";
        headerRange.format.font.bold = true;
        
        // Write sample data
        const dataRange = sheet.getRange("A2:I6");
        dataRange.values = sampleData;
        
        setMessage("Sample data loaded! Select rows 2-6 and click Predict.");
        
      });
    } catch (err) {
      setMessage("Error loading sample data: " + err.message);
    }
  };

  return (
    <div className="app" style={{ padding: "20px", fontFamily: "Segoe UI, sans-serif" }}>
      <h1 style={{ color: "#2b579a", marginBottom: "20px" }}>Employee Churn Predictor</h1>
      
      <div style={{ marginBottom: "20px" }}>
        <p style={{ fontSize: "14px", color: "#666", marginBottom: "10px" }}>
          This tool predicts employee churn probability using machine learning.
        </p>
        
        <div style={{ marginBottom: "15px" }}>
          <button 
            onClick={handleLoadSampleData}
            style={{
              backgroundColor: "#28a745",
              color: "white",
              border: "none",
              padding: "8px 16px",
              borderRadius: "4px",
              cursor: "pointer",
              marginRight: "10px"
            }}
          >
            Load Sample Data
          </button>
        </div>
        
        <div style={{ marginBottom: "15px" }}>
          <button 
            onClick={handlePredict}
            disabled={isLoading}
            style={{
              backgroundColor: isLoading ? "#6c757d" : "#007bff",
              color: "white",
              border: "none",
              padding: "10px 20px",
              borderRadius: "4px",
              cursor: isLoading ? "not-allowed" : "pointer",
              fontSize: "16px"
            }}
          >
            {isLoading ? "Processing..." : "Predict Churn"}
          </button>
        </div>
      </div>
      
      <div style={{ 
        padding: "10px", 
        backgroundColor: "#f8f9fa", 
        borderRadius: "4px",
        border: "1px solid #dee2e6"
      }}>
        <p style={{ margin: "0", fontSize: "14px" }}>{message}</p>
      </div>
      
      <div style={{ marginTop: "20px", fontSize: "12px", color: "#666" }}>
        <h4>Instructions:</h4>
        <ol style={{ paddingLeft: "20px" }}>
          <li>Click "Load Sample Data" to add example data</li>
          <li>Select the data rows (excluding header)</li>
          <li>Click "Predict Churn" to get predictions</li>
          <li>Results will appear in a new column</li>
        </ol>
        
        <h4>Required Data Format:</h4>
        <p>satisfaction_level, last_evaluation, number_project, average_montly_hours, time_spend_company, Work_accident, promotion_last_5years, Department, salary</p>
      </div>
    </div>
  );
};

export default App;
