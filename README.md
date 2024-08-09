# LinearProgramming
**Problem**

Imagine a company that manufactures products requiring different widths of paper. The company has large paper reels that need to be cut into specific widths for various products. Each product has a certain demand, and the goal is to optimize the cutting process to minimize waste while meeting the demand for each product.
The company has a total reel size of 2200 mm available. The challenge is to determine how many cuts of each width should be made from the reel to meet the demand for each SKU, while also ensuring that the total width used does not exceed the available reel size.
**Sample Data:**

sku	width_type	demand	length	gsm	month
SKU1	500	15000	1000	150	June
SKU2	600	12000	1200	170	June
SKU3	700	18000	1100	160	June
SKU4	800	13000	1150	155	June
SKU5	900	14000	1050	165	June

**Solution Using Linear Programming**

To solve this problem, we’ll use Linear Programming to optimize the cutting process. The approach involves the following steps:

Data Preparation: We begin by processing the data to calculate the demand in terms of the required width.
Initial Model Setup: We create a linear programming model to maximize the utilization of the reel size while satisfying the demand constraints.
Running the Model: We solve the model and adjust the demand based on the solution.
Iterative Optimization: The process is repeated until all demand is met or no further optimization is possible.
Generating the Final Output: We generate a final output that shows the optimum reel size and the total demand met for each SKU.

 ---Detailed explanation on the LPP problem is given in this [blog]((https://medium.com/p/36374463007e)) post ---

**User Interface is developed using plotly Dash with necessary functionalities like Edit, save, upload, Download and send email. Refer home_new.py**
