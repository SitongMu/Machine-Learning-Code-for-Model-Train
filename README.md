This project contains a seven step process containing data colllection, standardlization, train, test and auto train code.

1DataCollection is used to collect data from serial port and save them as an excel file.
2-3COmbin&Generate&Standardisation is used to deal with collected data. The first and last part of the data will be cut to avoid redundant influence. The data will be combined together and standardized to convince the train and test actions later.
4Train_together is used to train the model based on TensorFlow model and the data processed before.
5Realtimerun is used to run the function in real time.
6Testseperate is used to test the accuracy of the model with a seperate test file.
7Auto Train is used to check the difference with different combinations of the inputs.
