# To run this script, you must install the full version of OpenCV, which includes
# the dnn_superres module. Please run the following command in your terminal:
# pip uninstall opencv-python
# pip install opencv-contrib-python

import cv2
import numpy as np
import os
import tkinter as tk
from tkinter import filedialog

def enhance_image_quality(input_path, output_path, scale_factor=2):
    """
    Enhances the quality of an image using a Super-Resolution deep learning model.

    Args:
        input_path (str): The path to the input low-quality image.
        output_path (str): The path to save the enhanced output image.
        scale_factor (int): The factor by which to upscale the image. The model
                           used in this script is pre-trained for a 2x factor.
    """
    try:
        # Step 1: Read the image from the specified path
        image = cv2.imread(input_path)

        # Check if the image was loaded successfully
        if image is None:
            print(f"Error: Could not read the image from {input_path}. Please check the file path.")
            return

        print(f"Original image dimensions: {image.shape[1]}x{image.shape[0]}")

        # Step 2: Initialize the Super-Resolution model
        # You need to download the ESPCN_x2.pb file and place it in the same directory as the script.
        sr = cv2.dnn_superres.DnnSuperResImpl_create()
        model_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ESPCN_x2.pb')
        
        # Check if the model file exists
        if not os.path.exists(model_path):
            print("Error: Model file 'ESPCN_x2.pb' not found.")
            print("Please download it and place it in the same directory as this script.")
            return

        sr.readModel(model_path)
        sr.setModel("espcn", scale_factor)

        # Step 3: Apply the Super-Resolution model to upscale the image
        enhanced_image = sr.upsample(image)
        
        print(f"Enhanced image dimensions: {enhanced_image.shape[1]}x{enhanced_image.shape[0]}")

        # Step 4: Save the enhanced image to the output path
        # Using a .png extension to avoid further compression artifacts from JPEG
        cv2.imwrite(output_path, enhanced_image)

        print(f"Image successfully enhanced and saved to {output_path}")

    except Exception as e:
        print(f"An error occurred: {e}")

def open_file_dialog():
    """
    Opens a file dialog to allow the user to select an image file.
    """
    # Create a Tkinter root window
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Open the file dialog and get the selected file path
    file_path = filedialog.askopenfilename(
        title="Select a low-quality image file",
        filetypes=(("Image files", "*.jpg *.jpeg *.png"), ("All files", "*.*"))
    )

    # Destroy the root window after the file is selected
    root.destroy()
    
    return file_path

if __name__ == "__main__":
    # Get the input file path from the file dialog
    input_path = open_file_dialog()

    if input_path:
        # Get the directory and file name from the selected path
        input_directory, input_file_name = os.path.split(input_path)
        
        # Define the output file name
        output_file_name = 'enhanced_' + os.path.splitext(input_file_name)[0] + '.png'
        output_path = os.path.join(input_directory, output_file_name)

        # Enhance the image quality
        enhance_image_quality(input_path, output_path)
    else:
        print("No file was selected. The program will exit.")
