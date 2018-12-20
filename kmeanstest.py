from sklearn.cluster import KMeans
import numpy as np

ray = np.array([[2,2],[4,4],[16,16],[16,15]])
kmeans = KMeans(n_clusters=2)
# Fitting the input data
kmeans = kmeans.fit(ray)
# Getting the cluster labels
labels = kmeans.predict(ray)
# Centroid values
centroids = kmeans.cluster_centers_
print(centroids)
