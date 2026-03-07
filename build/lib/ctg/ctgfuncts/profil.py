
__all__ = ["trace_profil"]

import gpxpy
import gpxpy.gpx
import matplotlib.pyplot as plt
import numpy as np
from math import asin, cos, radians, sin, sqrt,acos

def distance_(phi1, lon1,phi2, lon2):

    """
    https://geodesie.ign.fr/contenu/fichiers/Distance_longitude_latitude.pdf
    """
    

    rad = 6371
    phi1, lon1 = radians(phi1), radians(lon1)
    phi2, lon2 = radians(phi2), radians(lon2)
    #dist = 2 * rad * asin(sqrt(sin((phi2 - phi1) / 2) ** 2
    #                         + cos(phi1) * cos(phi2) * sin((lon2 - lon1) / 2) ** 2))
    v = sin(phi1)*sin(phi2)+cos(phi1)*cos(phi2)*cos(lon1-lon2)
    if v>1 : v=1

    dist = 2*rad* acos(v)

    return dist


def trace_profil(gpx_path,mode="3D"):
    """Plot a 2D or 3D profile from a .gpx file
    """
    gpx_file = open(gpx_path, 'r')
    gpx = gpxpy.parse(gpx_file)
    offset = -100
    
    lat = []
    long = []
    h = []
    y = []
    d = []
    delta_d_list = []
    for track in gpx.tracks:
        print('---')
        for segment in track.segments:
            for idx,point in enumerate(segment.points):
                lat.append(point.latitude)
                long.append(point.longitude)
                
                if idx==0:
                    d.append(0)
                    if mode == "3D" :h.append(offset)
                    y.append(0)
                    h.append(point.elevation)
                    y.append(0)
                    
                else:
                    delta_d = distance_(lat[idx-1],long[idx-1],point.latitude,point.longitude)
                    delta_d_list.append(delta_d)
                    h.append(point.elevation)
                    y.append(-idx*0.001)
        y.append(y[idx+1])
        if mode == "3D" : h.append(offset)
        d = list(np.cumsum(delta_d_list))
        d = [0,d[0]] + d + [d[-1],0]
        y.append(0)
        if mode == "3D" : h.append(offset)
    
    if mode != "3D":
        fig = plt.figure() 
        ax = fig.add_subplot(111, projection='3d') 
        ax.plot(lat, long, h, color='blue') 
        plt.axis('off') 
    else:
        fig = plt.figure() 
        ax = fig.add_subplot(111, projection='3d') 
        ax.plot(d, y, h, color='blue') 
        plt.ylim(0,-20)
    
    plt.show()

 