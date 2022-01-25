import sys
sys.path.extend([r'C:\\Users\\hiepdang\\PycharmProjects\\DataAnalytics',
                 r'C:/Users/hiepdang/PycharmProjects/DataAnalytics'])

from post_phs import *

preproccessing.clearBreakeven()
post.breakeven(exchanges=['UPCOM'])
