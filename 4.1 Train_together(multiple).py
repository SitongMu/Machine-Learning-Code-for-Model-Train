import pandas as pd
import numpy as np
import tensorflow as tf
from tensorflow import keras
from sklearn.preprocessing import StandardScaler
from sklearn.utils import shuffle
from sklearn.model_selection import train_test_split

# Load data from Excel file
excel_file_path = r'C:\Users\2816624M\Desktop\lab2TX2b_tr.xlsx'
sheet_name = 'Standardized_Data'  # Replace with the name of your data sheet
data = pd.read_excel(excel_file_path, sheet_name='Standardized_Data')

data['V1'] = data['V1'].astype(float)
data['V2'] = data['V2'].astype(float)
data['V3'] = data['V3'].astype(float)
data['V4'] = data['V4'].astype(float)
data['V5'] = data['V5'].astype(float)
data['V6'] = data['V6'].astype(float)
data['V7'] = data['V7'].astype(float)
data['V8'] = data['V8'].astype(float)


data = data.dropna()  # Remove rows with missing values
data['Situation'] = data['Situation'].astype(int)  # Convert to integer

# Extract features (sensor readings) and labels
X = data[['V1', 'V2', 'V3', 'V4', 'V5', 'V6', 'V7', 'V8']].values
y = data['Situation'].values

# Shuffle the data
X, y = shuffle(X, y, random_state=42)

# One-hot encode the target labels
y = keras.utils.to_categorical(y, num_classes=11)

# Split the data into training and testing sets
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.1, random_state=42)

# Build a TensorFlow model
model = keras.Sequential([
    keras.layers.Input(shape=(8,)),  # Input layer with 10 features
    keras.layers.Dense(64, activation='relu'),  # Hidden layer with ReLU activation
    keras.layers.Dense(32, activation='relu'),  # Hidden layer with ReLU activation
    keras.layers.Dense(11, activation='sigmoid')  # Output layer with sigmoid activation (3 classes)
])

# Compile the model with categorical cross-entropy loss
model.compile(optimizer='adam', loss='categorical_crossentropy', metrics=['accuracy'])

# Set up callbacks for training (not included when saving the model)
callbacks = [
    keras.callbacks.EarlyStopping(patience=10, restore_best_weights=True)
]

# Train the model on training data
epochs = 50  # Increase the number of epochs for better convergence
batch_size = 32
model.fit(X_train, y_train, epochs=epochs, batch_size=batch_size, verbose=2, callbacks=callbacks, validation_split=0.1)

# Evaluate the model on the test data
test_loss, test_accuracy = model.evaluate(X_test, y_test)
print(f'Test loss: {test_loss:.4f}')
print(f'Test accuracy: {test_accuracy:.4f}')

model.save('horizontal1rmodel.keras')

# Save the split test data to an Excel file
# test_df = pd.DataFrame(data=X_test, columns=['V1', 'V2', 'V3', 'V4', 'V5', 'V6', 'V7', 'V8'])

# test_df['Situation'] = y_test.argmax(axis=1)
# test_df.to_excel(r'C:\Users\2816624M\Desktop\8rd1verticalfullte.xlsx', index=False)

# Predict category probabilities for test data
category_probabilities_test = model.predict(X_test)

# Assuming you have three categories (0, 1, 2)
num_categories = 11

# Create a DataFrame to store the results for the test data
result_test_df = pd.DataFrame({
    'Real_Category': y_test.argmax(axis=1),  # Real category (0, 1, 2) for test data
    'Predicted_Category': category_probabilities_test.argmax(axis=1),  # Predicted category (0, 1, 2) for test data
})

# Add predicted probabilities for each category to the DataFrame
for i in range(num_categories):
    result_test_df[f'Probability_Category_{i}'] = category_probabilities_test[:, i]

# Save the results for the test data to an Excel file
result_test_df.to_excel(r'C:\Users\2816624M\Desktop\horizontal_te.xlsx', index=False)
