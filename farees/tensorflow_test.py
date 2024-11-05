import tensorflow as tf

# Define a constant tensor
hello = tf.constant("Hello, World!")

# Create a TensorFlow session
with tf.compat.v1.Session() as sess:
    # Run the session to evaluate the tensor
    result = sess.run(hello)
    # Print the result
    print(result.decode())
