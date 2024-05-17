import pandas as pd
import numpy as np
import tensorflow as tf
from sklearn.metrics import confusion_matrix
import seaborn as sns
import matplotlib.pyplot as plt

# Modify the file path easily
excel_file_path = r'C:\Users\2816624M\Desktop\centertx_te.xlsx'  # Update with your test data file path
sheet_name = 'Standardized_Data'  # Replace with the name of your test data sheet
# Load test data from Excel file
test_data = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# Convert columns to float and drop missing values
sensor_columns = ['V1', 'V2', 'V3', 'V4', 'V5', 'V6', 'V7', 'V8']
test_data[sensor_columns] = test_data[sensor_columns].astype(float)
test_data = test_data.dropna()

# Define the zones and corresponding sensors (same as training)
zone_sensors = {
    # 1: ['V7', 'V8'],
    # 2: ['V5', 'V6'],
    # 3: ['V3', 'V4'],
    # 4: ['V1', 'V2']
    1: ['V1', 'V2'],  # Add other zones and their sensors here
    2: ['V3', 'V4'],
    3: ['V5', 'V6'],
    4: ['V7', 'V8']
}

# Initialize lists to store combined predictions and actual labels
combined_predictions = []
combined_actuals = []

# Iterate through each zone
for zone, sensors in zone_sensors.items():
    # Load the trained model for the current zone
    model = tf.keras.models.load_model(f'1r_{zone}_model.keras')

    # Filter and preprocess test data for the current zone
    zone_test_data = test_data.copy()
    zone_test_data['Situation'] = np.where(zone_test_data['Situation'] == zone, zone, 'Nopeo')
    zone_test_data['Situation'] = zone_test_data['Situation'].astype('category').cat.codes

    # Extract features and labels
    X_test = zone_test_data[sensors].values
    y_test = zone_test_data['Situation'].values
    y_test_encoded = tf.keras.utils.to_categorical(y_test)

    # Predict and evaluate
    test_loss, test_accuracy = model.evaluate(X_test, y_test_encoded)
    predictions = model.predict(X_test)
    predicted_classes = np.argmax(predictions, axis=1)

    # Aggregate predictions and actuals
    combined_predictions.extend(predicted_classes)
    combined_actuals.extend(y_test)

    print(f'Zone: {zone}, Test Accuracy: {test_accuracy:.4f}')

# Generate and save the combined confusion matrix
cm = confusion_matrix(combined_actuals, combined_predictions)
plt.figure(figsize=(10, 7))
sns.heatmap(cm, annot=True, fmt='d')
plt.title('Combined Confusion Matrix for All Zones')
plt.ylabel('Actual')
plt.xlabel('Predicted')
plt.savefig(r'C:\Users\2816624M\Desktop\combined_confusion_matrix.png')
