import serial
import tensorflow as tf
import numpy as np
from sklearn.preprocessing import StandardScaler

# Load the trained model
model = tf.keras.models.load_model('horizontal1rmodel.keras')

# Setup serial connection (adjust 'COM3' and baudrate as per your configuration)
arduino = serial.Serial('COM4', 9600, timeout=1)

results = []  # List to store results
counter = 0  # Counter for keeping track of results

while True:
    # Read data from Arduino
    line = arduino.readline()
    if line:
        data_str = line.decode('utf-8').strip()  # Convert bytes to string and strip newlines
        values = data_str.split(',')  # Split string into list based on comma
        if len(values) == 8:  # Ensure there are 8 sensor readings
            try:
                data = np.array([float(v) for v in values])  # Convert string values to float
                data = data.reshape(1, -1)  # Reshape for the model

                # Example: Normalize or standardize data if your model requires
                # data = scaler.transform(data)

                prediction = model.predict(data)
                category = np.argmax(prediction, axis=1)[0]

                # Append result to results list
                if category == 0:
                    results.append("Nopeo")
                elif category == 1:
                    results.append("Zone1")
                elif category == 2:
                    results.append("Zone2")
                elif category == 3:
                    results.append("Zone3")
                elif category == 4:
                    results.append("Zone4")

                counter += 1

                # Print and reset every 10 readings
                if counter == 1:
                    print(", ".join(results))
                    results = []
                    counter = 0

            except ValueError:
                print("Invalid data received")
        else:
            print("Incomplete data received")
