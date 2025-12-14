import tensorflow as tf

x = tf.constant([1, 2, 3])
y = tf.constant([4, 5, 6])

result = tf.multiply(x, y)

with tf.version() as sess:
    output = sess.run(result)
    print(output)
