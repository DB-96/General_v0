# Import necessary libraries
import numpy as np
from numba import jit
import time

# Load the input data
# Replace with raw data files
data = np.load('ct_scan_data.npy')

# Define reconstruction parameters
angles = np.linspace(0., 180., max(data.shape), endpoint=False)
num_angles = len(angles)
sinogram_shape = (num_angles, data.shape[1])

# Define the backprojection calculation
@jit(nopython=True)
def calculate_backprojection(sinogram, angles, output_shape):
    backprojection = np.zeros(output_shape, dtype=sinogram.dtype)
    for i, angle in enumerate(angles):
        for j in range(sinogram.shape[1]):
            x, y = np.dot([[np.cos(angle), np.sin(angle)], [-np.sin(angle), np.cos(angle)]],
                          [j - sinogram.shape[1] / 2, i - sinogram.shape[0] / 2])
            x += output_shape[1] / 2
            y += output_shape[0] / 2
            if 0 <= x < output_shape[1] and 0 <= y < output_shape[0]:
                backprojection[int(y), int(x)] += sinogram[i, j]
    return backprojection

# Calculate the sinogram
sinogram = calculate_sinogram(data, angles, sinogram_shape)

# Start the timer
start_time = time.time()

# Calculate the backprojection
reconstruction = calculate_backprojection(sinogram, angles, data.shape)

# End the timer
end_time = time.time()

# Calculate the runtime
runtime = end_time - start_time

# Print the runtime
print(f"Runtime: {runtime} seconds")

# Save the output
np.save('reconstruction_bp.npy', reconstruction)


# angles = np.linspace(0., 180., max(data.shape), endpoint=False)
# num_angles = len(angles)
# sinogram_shape = (num_angles, data.shape[1])

# # Define the sinogram calculation
# @jit(nopython=True)
# def calculate_sinogram(data, angles, sinogram_shape):
#     sinogram = np.empty(sinogram_shape, dtype=data.dtype)
#     for i, angle in enumerate(angles):
#         rotation_matrix = np.array([[np.cos(angle), -np.sin(angle)],
#                                     [np.sin(angle), np.cos(angle)]])
#         for j in range(sinogram_shape[1]):
#             x, y = np.dot(rotation_matrix, [j - sinogram_shape[1] / 2, 0])
#             x += data.shape[1] / 2
#             y += data.shape[0] / 2
#             sinogram[i, j] = map_coordinates(data, [y, x], order=1, mode='nearest')
#     return sinogram

# # Calculate the sinogram
# sinogram = calculate_sinogram(data, angles, sinogram_shape)

# # Define the filtered backprojection
# @jit(nopython=True)
# def filtered_backprojection(sinogram, angles, output_shape):
#     reconstruction = np.zeros(output_shape, dtype=sinogram.dtype)
#     for i, angle in enumerate(angles):
#         rotation_matrix = np.array([[np.cos(angle), -np.sin(angle)],
#                                     [np.sin(angle), np.cos(angle)]])
#         for j in range(output_shape[1]):
#             x, y = np.dot(rotation_matrix, [j - output_shape[1] / 2, 0])
#             x += sinogram.shape[1] / 2
#             y += sinogram.shape[0] / 2
#             reconstruction[j, :] += map_coordinates(sinogram[i, :], [y, x], order=1, mode='nearest')
#     return reconstruction

# # Reconstruct the image
# reconstruction = filtered_backprojection(sinogram, angles, data.shape)

# # Save the output
# np.save('reconstruction.npy', reconstruction)