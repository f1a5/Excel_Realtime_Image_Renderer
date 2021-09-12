# Excel_Realtime_Image_Renderer
Rendering of jpg/png images on an Excel Sheet using the individual cells of the sheet as pixels. It loads the target image file into memory, reads the color values of each pixel into a numpy array and then it fills the cells of the excel sheet corresponding to the indexes of numpy array.

## Screenshots
![image](https://user-images.githubusercontent.com/73558085/132998138-6f4c6791-9e7b-41f9-9681-bf919b1b45a7.png)
![image](https://user-images.githubusercontent.com/73558085/132998154-c6d32461-d09d-4c39-9feb-d3a892889229.png)
![image](https://user-images.githubusercontent.com/73558085/132998165-8c4629cb-b831-4ce8-942f-cc43281cc96d.png)

## To Use
Clone the repo and install the required dependencies (virtual environment is recommended) by running ```pip install -r requirements.txt```. Then open the *Configuration.py* and edit the Paths to input and output files before running *Renderer.py*.
To adjust cell sizes, do ```ctrl+A``` and press format cells button from tool bar. A row height of 0.8 and column length value of 0.2 sets the cells size as 1x1 pixels.

