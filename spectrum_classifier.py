# for reading data
import keras.callbacks
import pandas as pd
import numpy as np
from sklearn.preprocessing import LabelEncoder

import keras
from keras.wrappers.scikit_learn import KerasClassifier
from keras.utils import np_utils
from keras.utils.vis_utils import plot_model

# for modeling
from keras.models import Sequential
from keras.layers import Dense, Dropout
from keras.callbacks import EarlyStopping

from sklearn.metrics import confusion_matrix
from sklearn.metrics import classification_report


import matplotlib.pyplot as plt


df = pd.read_excel("optical_spectrum_dataset.xlsx", engine='openpyxl')
df = df.sample(frac = 1)



# shuffle the dataset!
#df = df.sample(frac=1).reset_index(drop=True)

# split into X and Y


Y = df['element']
X = df.drop(['element'], axis=1)
X = np.array(X)

# work with labels
# encode class values as integers
encoder = LabelEncoder()
encoder.fit(Y)
encoded_Y = encoder.transform(Y)
# convert integers to dummy variables (i.e. one hot encoded)
class_y = np_utils.to_categorical(encoded_Y)


model = Sequential()
model.add(Dense(10, input_shape=(X.shape[1],), activation='relu'))
model.add(Dense(10, input_shape=(X.shape[1],), activation='relu'))
model.add(Dense(3, activation='softmax'))
model.summary()

model.compile(optimizer='rmsprop',
              loss='categorical_crossentropy',
              metrics=['accuracy']
              )

#get block diagram of neural network model
#plot_model(model, to_file="my_model.png", show_shapes=True)

# early stopping callback
# This callback will stop the training when there is no improvement in
# the validation loss for 10 consecutive epochs.
es = keras.callbacks.EarlyStopping(monitor='val_loss',
                                   mode='min',
                                   patience=10,
                                   restore_best_weights=True)

# now we just update our model fit call
history = model.fit(X,
                    class_y,
                    callbacks=[es],
                    epochs=1000,
                    batch_size=10,
                    shuffle=True,
                    validation_split=0.2,
                    verbose=1)

history_dict = history.history
#learning curve
#accuracy
acc = history_dict['accuracy']
val_acc = history_dict['val_accuracy']
#range of X (no. of epochs
epochs = range(1, len(acc)+1)

#plot
plt.plot(epochs, acc, 'r', label='Training accuracy')
plt.plot(epochs, val_acc, 'b', label='Validation accuracy')
plt.title("Training and validation ccuracy")
plt.xlabel('Epochs')
plt.ylabel('Accuracy')
plt.legend()
plt.show()


#prediction

preds = model.predict(X)
#pred[0] will contain prediction all classes
#and the sum of all classes for single prediction should be equal to 1
print(np.sum(preds[0]))


print(confusion_matrix(class_y.argmax(axis=1), preds.argmax(axis=1)))
print(classification_report(class_y.argmax(axis=1), preds.argmax(axis=1)))


#saving modle
#then loading it
model_name = "spectrum_model.h5"
model.save(model_name)
model = keras.models.load_model(model_name)


#single prediction
#need to expand that is to convert array shape (1, 5) to (5, 1) for single value
#need array [[value1, value2, value3, value4, value5]]
no_of_classes = len(encoder.classes_)

x = np.expand_dims(X[22], axis=0)
new_pred = model.predict(x)
max_value = max(new_pred[0])
for i in range(no_of_classes):
    if new_pred[0][i] == max_value:
        index = i

predicted_class = encoder.classes_[index]
print('predicted Class = ', predicted_class)



