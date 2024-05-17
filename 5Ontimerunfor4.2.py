import serial
import tensorflow as tf
import numpy as np

# Load the trained models for each pair of sensors
models = {
    'model_12': tf.keras.models.load_model('1r_4_model.keras'),  # Monitors Zone 4
    'model_34': tf.keras.models.load_model('1r_3_model.keras'),  # Monitors Zone 3
    'model_56': tf.keras.models.load_model('1r_2_model.keras'),  # Monitors Zone 2
    'model_78': tf.keras.models.load_model('1r_1_model.keras')   # Monitors Zone 1
}

# Setup serial connection (adjust 'COM4' and baudrate as per your configuration)
arduino = serial.Serial('COM4', 9600, timeout=1)

while True:
    # Read data from Arduino
    line = arduino.readline()
    if line:
        data_str = line.decode('utf-8').strip()  # Convert bytes to string and strip newlines
        values = data_str.split(',')  # Split string into list based on comma
        if len(values) == 8:  # Ensure there are 8 sensor readings
            try:
                for i in range(0, 8, 2):
                    model_key = f'model_{i+1}{i+2}'
                    data_pair = np.array([float(values[i]), float(values[i+1])]).reshape(1, -1)
                    
                    # If your models expect standardized or normalized data, apply the same here
                    # Example: data_pair = scaler.transform(data_pair)
                    
                    prediction = models[model_key].predict(data_pair)
                    category = np.argmax(prediction, axis=1)[0]

                    # Adjusting zone numbering according to sensor pair
                    zone = 4 - (i // 2) 
                    if category == 0:
                        print(f"No people in Zone {zone}")
                    else:
                        print(f"Someone in Zone {zone}")
            except ValueError:
                print("Invalid data received")
        else:
            print("Incomplete data received")
