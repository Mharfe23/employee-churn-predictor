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
          
          // Check if prediction columns already exist
          const usedRange = sheet.getUsedRange();
          usedRange.load("columnCount, values");
          await context.sync();
          
          let predictionColumn, yesNoColumn;
          
          // Check if "Churn Probability" and "Will Leave?" columns already exist
          const headers = usedRange.values[0] || [];
          const churnProbIndex = headers.findIndex(header => header === "Churn Probability");
          const willLeaveIndex = headers.findIndex(header => header === "Will Leave?");
          
          if (churnProbIndex !== -1 && willLeaveIndex !== -1) {
            // Use existing columns
            predictionColumn = String.fromCharCode(65 + churnProbIndex);
            yesNoColumn = String.fromCharCode(65 + willLeaveIndex);
          } else {
            // Create new columns
            const nextColumn = usedRange.columnCount + 1;
            predictionColumn = String.fromCharCode(64 + nextColumn);
            yesNoColumn = String.fromCharCode(65 + nextColumn);
          }
          
          // Create range for predictions (same number of rows as selected data, starting from row 2)
          const predictionRange = sheet.getRange(`${predictionColumn}2:${predictionColumn}${range.values.length + 1}`);
          const yesNoRange = sheet.getRange(`${yesNoColumn}2:${yesNoColumn}${range.values.length + 1}`);
          
          // Format predictions as percentages and write to Excel
          const formattedPredictions = predictions.map(pred => {
            const churnProbability = pred[1]; // Probability of leaving (churn)
            return [churnProbability.toFixed(4)]; // Format as 4 decimal places
          });
          
          // Create Yes/No predictions based on threshold (0.5 = 50%)
          const threshold = 0.5;
          const yesNoPredictions = predictions.map(pred => {
            const churnProbability = pred[1]; // Probability of leaving (churn)
            return [churnProbability >= threshold ? "Yes" : "No"];
          });
          
          predictionRange.values = formattedPredictions;
          yesNoRange.values = yesNoPredictions;
          
          // Add headers for both columns
          const predictionHeaderRange = sheet.getRange(`${predictionColumn}1`);
          const yesNoHeaderRange = sheet.getRange(`${yesNoColumn}1`);
          
          predictionHeaderRange.values = [["Churn Probability"]];
          yesNoHeaderRange.values = [["Will Leave?"]];
          
          // Format the headers
          predictionHeaderRange.format.fill.color = "#4472C4";
          predictionHeaderRange.format.font.color = "white";
          predictionHeaderRange.format.font.bold = true;
          
          yesNoHeaderRange.format.fill.color = "#C5504B";
          yesNoHeaderRange.format.font.color = "white";
          yesNoHeaderRange.format.font.bold = true;
          
          // Format prediction values as percentages
          predictionRange.format.numberFormat = "0.00%";
          
          // Format Yes/No column with conditional formatting
          yesNoRange.format.font.bold = true;
          
          if (churnProbIndex !== -1 && willLeaveIndex !== -1) {
            setMessage(`Predictions updated in columns ${predictionColumn} and ${yesNoColumn}!`);
          } else {
            setMessage(`Predictions added in columns ${predictionColumn} and ${yesNoColumn}!`);
          }
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
        
        // Sample data rows - diverse examples representing different scenarios
        const sampleData = [
          // High churn risk examples
          [0.11, 0.88, 7, 272, 4, 0, 0, "technical", "medium"],  // Low satisfaction, high evaluation, many projects
          [0.15, 0.92, 6, 280, 5, 0, 0, "sales", "low"],         // Very low satisfaction, high evaluation
          [0.25, 0.85, 8, 265, 6, 0, 0, "IT", "medium"],         // Low satisfaction, many projects, long hours
          
          // Medium churn risk examples
          [0.38, 0.53, 2, 157, 3, 0, 0, "sales", "low"],         // Low satisfaction, low evaluation
          [0.45, 0.67, 4, 200, 4, 0, 0, "support", "medium"],    // Medium satisfaction, medium evaluation
          [0.52, 0.78, 5, 220, 5, 0, 0, "marketing", "high"],    // Medium satisfaction, good evaluation
          
          // Low churn risk examples
          [0.72, 0.87, 5, 223, 5, 0, 0, "support", "low"],       // High satisfaction, good evaluation
          [0.85, 0.91, 4, 180, 3, 1, 1, "management", "high"],   // Very high satisfaction, promoted recently
          [0.78, 0.82, 3, 160, 2, 0, 1, "hr", "medium"],         // High satisfaction, promoted, few projects
          [0.90, 0.95, 6, 200, 4, 1, 0, "product_mng", "high"],  // Very high satisfaction, high evaluation
          
          // Edge cases
          [0.30, 0.45, 1, 120, 1, 0, 0, "accounting", "low"],    // New employee, low satisfaction
          [0.95, 0.98, 7, 250, 8, 1, 1, "RandD", "high"],        // Veteran employee, very satisfied
          [0.20, 0.75, 9, 300, 7, 0, 0, "technical", "medium"]   // Overworked, low satisfaction
        ];
        
        // Write headers
        const headerRange = sheet.getRange("A1:I1");
        headerRange.values = headers;
        headerRange.format.fill.color = "#4472C4";
        headerRange.format.font.color = "white";
        headerRange.format.font.bold = true;
        
        // Write sample data
        const dataRange = sheet.getRange("A2:I14");
        dataRange.values = sampleData;
        
        setMessage("Sample data loaded! Select rows 2-14 and click Predict.");
        
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
