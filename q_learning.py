#_*_encoding:utf-8_*_#
import numpy as np
GAMMA = 0.8
Q = np.zeros((6,6))
R=np.asarray([[-1,-1,-1,-1,0,-1],
   [-1,-1,-1,0,-1,100],
   [-1,-1,-1,0,-1,-1],
   [-1,0, 0, -1,0,-1],
   [0,-1,-1,0,-1,100],
   [-1,0,-1,-1,0,100]])
def getMaxQ(state):
    return max(Q[state, :])
def QLearning(state):
    curAction = None
    for action in xrange(6):
        if(R[state][action] == -1):
            Q[state, action]=0
        else:
            curAction = action
            Q[state,action]=R[state][action]+GAMMA * getMaxQ(curAction)
count=0
while count<100000:
    for i in xrange(6):
        QLearning(i)
    count+=1
print Q/5