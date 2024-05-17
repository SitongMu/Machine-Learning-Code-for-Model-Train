import pandas as pd
import numpy as np
import tensorflow as tf
from tensorflow import keras
from sklearn.utils import shuffle
from itertools import combinations


num_columns_to_choose =8# Update this with your preferred number
Trainfile_name = 'Lab2TXlonger_tr'
Testfile_name = 'Lab2TXlonger_te_combined'

# Function to build and train the model
# Function to build and train the model
def build_and_train_model(X_train, y_train, input_shape):
    model = keras.Sequential([
        keras.layers.Input(shape=(input_shape,)),
        keras.layers.Dense(64, activation='relu'),
        keras.layers.Dense(32, activation='relu'),
        keras.layers.Dense(5, activation='softmax')
    ])

    model.compile(optimizer='adam', loss='categorical_crossentropy', metrics=['accuracy'])
    callbacks = [keras.callbacks.EarlyStopping(patience=10, restore_best_weights=True)]
    model.fit(X_train, y_train, epochs=50, batch_size=32, verbose=0, callbacks=callbacks, validation_split=0.2)
    return model

def predict_and_store_results(model, X_test, y_test, filename):
    predicted = model.predict(X_test)
    predicted_labels = predicted.argmax(axis=1)
    real_labels = y_test.argmax(axis=1)

    results_df = pd.DataFrame({'Real': real_labels, 'Predicted': predicted_labels})
    result_excel_path = f'C:\\Users\\2816624M\\Desktop\\{Testfile_name}Auto\\{filename}.xlsx'
    results_df.to_excel(result_excel_path, index=False)




# Load training data
excel_file_path = f'C:/Users/2816624M/Desktop/{Trainfile_name}.xlsx'  # Update with your training file path
data = pd.read_excel(excel_file_path, sheet_name='Standardized_Data')

# Load external test data
test_file_path = f'C:/Users/2816624M/Desktop/{Testfile_name}.xlsx'  # Update this with the path to your test file
test_data = pd.read_excel(test_file_path, sheet_name='Standardized_Data')

# Define all columns and categories
all_columns = ['V1', 'V2', 'V3', 'V4','V5','V6','V7','V8']  # Update with all your column names


# Preprocess training data
data = data.dropna()
data['Situation'] = data['Situation'].astype(int)#.apply(update_category, args=(selected_categories,))

# Preprocess test data
test_data = test_data.dropna()
test_data['Situation'] = test_data['Situation'].astype(int)#.apply(update_category, args=(selected_categories,))

# Store results
results = []

# Iterate over all combinations of columns
for cols in combinations(all_columns, num_columns_to_choose):
    # Prepare training data for model
    X_train = data[list(cols)].values
    y_train = data['Situation'].values
    X_train, y_train = shuffle(X_train, y_train, random_state=42)
    y_train = keras.utils.to_categorical(y_train)
    

    # Prepare test data for model
    X_test = test_data[list(cols)].values
    y_test = test_data['Situation'].values
    #X_test, y_test = shuffle(X_test, y_test, random_state=42)
    y_test = keras.utils.to_categorical(y_test)

    # Build and train the model
    model = build_and_train_model(X_train,y_train, len(cols))

    _,accuracy = model.evaluate(X_test, y_test)

    # Store the results
    results.append((cols, accuracy))
    predict_and_store_results(model, X_test, y_test, f'Results_{cols}')

# Output the results
for combination, accuracy in results:
    print(f'|{combination}|{accuracy:.4f}|')
