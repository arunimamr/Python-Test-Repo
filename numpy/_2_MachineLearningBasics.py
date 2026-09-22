
import numpy
import matplotlib
import matplotlib.pyplot as plt
from scipy import stats

speed =  [99,86,87,88,111,86,103,87,94,78,77,85,86]

#Mean - The mean value is the average value --> mean()
speed_mean = numpy.mean(speed)
print(speed_mean)

#Median - The median value is the value in the middle, after you have sorted all the values --> median()
speed_median = numpy.median(speed)
print(speed_median)


#Mode - The Mode value is the value that appears the most number of times --> stats.mode()
speed_mode = stats.mode(speed)
print(speed_mode)


#Standard deviation -> Deviation from average/mean --> std()
speed = [32,111,138,28,59,77,97]
standard_deviation_val = numpy.std(speed)
print(standard_deviation_val)

#Variance ->< Another way of representing deviation --> var()
speed_variance = numpy.var(speed)
print(speed_variance)

speed_standard_deviation = numpy.sqrt(speed_variance)
print(speed_standard_deviation)

#MatPlotLib -> graphical representation

print("version of matplotlib - ", matplotlib.__version__)

xPoints = numpy.array([0, 6])
yPoints = numpy.array([0, 250])
plt.plot(xpoints, ypoints)
plt.show()