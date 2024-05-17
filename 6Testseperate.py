import pandas as pd
from tensorflow import keras
import numpy as np
import matplotlib.pyplot as plt
from sklearn.metrics import confusion_matrix
import seaborn as sns

# Modify this variable to change the file path easily
test_data_file_path = r'C:\Users\Timeline_1.xlsx'  # Update this path

# Load the new test data
test_data = pd.read_excel(test_data_file_path, sheet_name='Standardized_Data')
test_data = test_data.dropna()

# Convert columns to float and 'Situation' to int
test_columns = ['V1', 'V2', 'V3', 'V4', 'V5', 'V6', 'V7', 'V8']
for col in test_columns:
    test_data[col] = test_data[col].astype(float)
test_data['Situation'] = test_data['Situation'].astype(int)

# Extract features and labels
X_new_test = test_data[test_columns].values
y_new_test = test_data['Situation'].values

# One-hot encode the target labels
y_new_test = keras.utils.to_categorical(y_new_test, num_classes=5)

# Load the trained model
model = keras.models.load_model('horizontal1rmodel.keras')

# Predict the labels for the new test data
predicted_labels = model.predict(X_new_test)

# Evaluate the model accuracy
_, test_accuracy = model.evaluate(X_new_test, y_new_test)
print(f'New Test accuracy: {test_accuracy:.4f}')

cm = confusion_matrix(y_new_test.argmax(axis=1), predicted_labels.argmax(axis=1))
cm_normalized = cm.astype('float') / cm.sum(axis=1)[:, np.newaxis]

# Plotting the confusion matrix without automatic annotations
plt.figure(figsize=(10, 9))
sns.heatmap(cm_normalized, annot=False, fmt=".2f", cmap='GnBu')  # Set annot=False here
plt.xlabel('Predicted', fontsize=20, fontname='Times New Roman')
plt.ylabel('True', fontsize=20, fontname='Times New Roman')

# Setting custom tick labels
class_labels = ['Nope', 'Zone1', 'Zone2', 'Zone3', 'Zone4']  # replace with your class names
plt.xticks(np.arange(len(class_labels)) + 0.5, class_labels, fontsize=14, rotation=45)
plt.yticks(np.arange(len(class_labels)) + 0.5, class_labels, fontsize=14, rotation=45)

# Manually add annotations with adjusted font size
annotation_fontsize = 16  # Adjust this value as needed
for i in range(cm_normalized.shape[0]):
    for j in range(cm_normalized.shape[1]):
        plt.text(j + 0.5, i + 0.5, f'{cm_normalized[i, j]:.2f}',
                 horizontalalignment='center',
                 verticalalignment='center',
                 fontsize=annotation_fontsize,  # Set the fontsize here
                 color="white" if cm_normalized[i, j] > cm_normalized.max()/2. else "black")

# Save the confusion matrix as a heatmap image
plt.savefig(r'C:\Users\confusion_matrix_1.png', format='png')
plt.show()

# Predict the labels for the new test data
predicted_labels = model.predict(X_new_test)

# Transform predicted labels from one-hot encoded to original form
predicted_labels_original = np.argmax(predicted_labels, axis=1)

# Create a DataFrame with original and predicted labels
output_df = pd.DataFrame({
    'Original Category': np.argmax(y_new_test, axis=1),
    'Predicted Category': predicted_labels_original
})

# Save the DataFrame to an Excel file
output_file_path = r'C:\Users\predictions.xlsx'  # Specify your desired file path
output_df.to_excel(output_file_path, index=False)
