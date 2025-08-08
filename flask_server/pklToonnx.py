import xgboost as xgb
import numpy as np
from onnxmltools.convert.xgboost import convert

# Load your XGBoost model
model = xgb.Booster()
model.load_model('XGBoost.pkl')

# Convert to ONNX
onnx_model = convert(model, name='xgboost', initial_types=[('input', np.float32, [None, 9])])

# Save the model
with open("xgboost.onnx", "wb") as f:
    f.write(onnx_model.SerializeToString())