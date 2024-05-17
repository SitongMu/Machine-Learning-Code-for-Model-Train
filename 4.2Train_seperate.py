import pandas as pd
import numpy as np
import tensorflow as tf
from tensorflow import keras
from sklearn.model_selection import train_test_split
from sklearn.utils import shuffle

# Load data from Excel file
excel_file_path = r'C:\Users\2816624M\Desktop\centertx_tr.xlsx'  # Update with your file path
sheet_name = 'Standardized_Data'  # Replace with the name of your data sheet
data = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# Convert columns to float
sensor_columns = ['V1', 'V2', 'V3', 'V4', 'V5', 'V6', 'V7', 'V8']
for col in sensor_columns:
    data[col] = data[col].astype(float)

# Drop rows with missing values
data = data.dropna()

# Define the zones and corresponding sensors
zone_sensors = {
    # 1: ['V1', 'V2'],  # Add other zones and their sensors here
    # 2: ['V3', 'V4'],
    # 3: ['V5', 'V6'],
    # 4: ['V7', 'V8']
    1: ['V1', 'V2'],  # Add other zones and their sensors here
    2: ['V3', 'V4'],
    3: ['V5', 'V6'],
    4: ['V7', 'V8']
}

for zone, sensors in zone_sensors.items():
    # Filter data for current zone
    zone_data = data.copy()
    zone_data['Situation'] = np.where(zone_data['Situation'] == zone, zone, 'Nopeo')
    zone_data['Situation'] = zone_data['Situation'].astype('category').cat.codes

    # Extract features and labels for the current zone
    X = zone_data[sensors].values
    y = zone_data['Situation'].values

    # One-hot encode the target labels
    y = keras.utils.to_categorical(y)

    # Shuffle the data
    X, y = shuffle(X, y, random_state=42)

    # Split the data into training and testing sets
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

    # Define the model
    model = keras.Sequential([
        keras.layers.Input(shape=(len(sensors),)),  # Adjust input shape based on sensors
        keras.layers.Dense(64, activation='relu'),
        keras.layers.Dense(32, activation='relu'),
        keras.layers.Dense(y.shape[1], activation='softmax')  # Adjust output layer based on classes
    ])

    # Compile the model
    model.compile(optimizer='adam', loss='categorical_crossentropy', metrics=['accuracy'])

    # Train the model
    model.fit(X_train, y_train, epochs=50, batch_size=32, verbose=2, validation_split=0.1)

    # Evaluate the model
    test_loss, test_accuracy = model.evaluate(X_test, y_test)
    print(f'Zone: {zone}, Test loss: {test_loss:.4f}, Test accuracy: {test_accuracy:.4f}')

    # Save the model
    model.save(f'1r_{zone}_model.keras')
