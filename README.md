# heatmap-plotter
Excel VBA Macro to plot points onto a resizable range

##Instructions
To start, copy and paste your data into the second sheet of this workbook ('Data'). Then under configuration, enter the data ranges that you want to use for the x & y axis's and for the labels. You can use table notation or standard notation. 

Next, set the scale for your data set. You need to put a minimum and maximum value for both the x and y axis's. For example, a risk assessment might use the values 1 and 5. Alternatively, you can use the auto set button to automatically find the min and max values in your data set (this probably isn't the best scale for your data).

The red rectangle to the right is the plotting area. You can drag this to any dimension you need - the scale will automatically update to fit the red rectangle. If the red rectangle is not visible hit the 'New Plot' button. An error will occur if the red rectangle is not visible.

To create the plot hit the Create Points button.

Other than color, all circle attributes (size, shadow, etc), are based on the default circle shape. To edit this draw a circle, set it how you like, and right click->set as default shape
