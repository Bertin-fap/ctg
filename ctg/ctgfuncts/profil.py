
__all__ = ["trace_profil"]

import gpxpy
import gpxpy.gpx
import matplotlib.pyplot as plt
from math import asin, cos, radians, sin, sqrt,acos

def distance_(phi1, lon1,phi2, lon2):

    rad = 6371
    phi1, lon1 = radians(phi1), radians(lon1)
    phi2, lon2 = radians(phi2), radians(lon2)
    #dist = 2 * rad * asin(sqrt(sin((phi2 - phi1) / 2) ** 2
    #                         + cos(phi1) * cos(phi2) * sin((lon2 - lon1) / 2) ** 2))
    dist = 2*rad* acos(sin(phi1)*sin(phi2)+cos(phi1)*cos(phi2)*cos(lon1-lon2))
    return dist


def trace_profil(gpx_path,mode="3D"):
    """Plot a 2D or 3D profile from a .gpx file
    """
    gpx_file = open(gpx_path, 'r')
    gpx = gpxpy.parse(gpx_file)
    lat = []
    long = []
    h = []
    d = []
    y = []
  
    for track in gpx.tracks:
          for segment in track.segments:
            for idx,point in enumerate(segment.points):
                lat.append(point.latitude)
                long.append(point.longitude)
                
                if idx==0:
                    d.append(0)
                    if mode == "2D" :h.append(300)
                    y.append(0)
                    d.append(0)
                    h.append(point.elevation)
                    y.append(0)
                    
                else:
                    d.append(d[idx-1] + distance_(lat[idx-1],long[idx-1],point.latitude,point.longitude))
                    h.append(point.elevation)
                    y.append(-idx*0.001)
        d.append(d[idx+1])
        y.append(y[idx+1])
        if mode == "2D" : h.append(300)
        d.append(0)
        y.append(0)
        if mode == "2D" : h.append(300)
    
    if mode != "2D":
        fig = plt.figure() 
        plt.plot(d, h, color='blue') 
    else:
        fig = plt.figure() 
        ax = fig.add_subplot(111, projection='3d') 
        ax.plot(d, y, h, color='blue') 
        plt.ylim(0,-20)
    
    plt.show()
 