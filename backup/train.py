import numpy as np
import pandas as pd
from keras.layers import LSTM, Dense, Dropout
from keras.models import Sequential
from sklearn.preprocessing import MinMaxScaler

import config as c

df=pd.read_csv(c.data_file) 
sc = MinMaxScaler(feature_range = (0, 1))
training_set = df.iloc[:c.train_size, c.data_columns].values

print(df.head(5))
print(df.iloc[:c.train_size, c.data_columns].head(5))

test_set = df.iloc[c.train_size:, c.data_columns].values 
training_set_scaled = sc.fit_transform(training_set)

print(training_set_scaled.shape)

x_train = []
y_train = []
for i in range(c.lookback, c.train_size):
    x_train.append(training_set_scaled[i-c.lookback:i, 0])
    y_train.append(training_set_scaled[i, 0])

x_train, y_train = np.array(x_train), np.array(y_train)
x_train = np.reshape(x_train, (x_train.shape[0], x_train.shape[1], 1)) #(values, lookback, 1d output)

print(x_train.shape)
print(y_train.shape)

# model = Sequential()
# model.add(LSTM(units = 50, return_sequences = True, input_shape = (x_train.shape[1], 1)))
# model.add(Dropout(0.2))
# model.add(LSTM(units = 50, return_sequences = True))
# model.add(Dropout(0.2))
# model.add(LSTM(units = 50, return_sequences = True))
# model.add(Dropout(0.2))
# model.add(LSTM(units = 50))
# model.add(Dropout(0.2))
# model.add(Dense(units = 1))
# model.compile(optimizer = 'adam', loss = 'mean_squared_error')
# model.fit(x_train, y_train, epochs = c.epochs, batch_size = c.batch_size)
# model.save(c.model_file)
