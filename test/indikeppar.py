import pandas as pd
import matplotlib.pyplot as plt
import math

class Indikeppar:
    def __init__(self, equity_1_file, equity_2_file, limit=0):
        self.e1_data = pd.read_csv(equity_1_file)
        self.e2_data = pd.read_csv(equity_2_file)
        self.outdata = pd.read_csv(equity_1_file)
        self.outdata = self.outdata[['Time']]
        # Limit input to limit
        if limit != 0:
            self.e1_data = self.e1_data[-limit:]
            self.e2_data = self.e2_data[-limit:]
            self.outdata = self.outdata[-limit:]

    def calculateP(self, type_, idx, T):
        P = 1
        if type_ == 'A':
            for i in range(T+1):
                P *= (1 + self.e1_data.loc[idx - i, 'R'])
        else:
            for i in range(T+1):
                P *= (1 + self.e2_data.loc[idx - i, 'R'])
        return math.log(P)

    def calculateS(self, idx, T):
        S = self.calculateP('A', idx, T) - self.calculateP('B', idx, T)
        return S

    def calculateMiu(self, idx, t, T):
        miu = 0
        for i in range(t, T+1):
            miu += self.calculateS(idx, i) / (T-t+1)
        return miu

    def calculateSigma(self, idx, t, T):
        miu = self.calculateMiu(idx, t, t)
        sigma = 0
        for i in range(t, T+1):
            sigma += (self.calculateS(idx, i) - miu) / (T-t+1)
        return sigma

    def calculateTtest(self, idx, t, T):
        tTest = (self.calculateS(idx, T) - self.calculateMiu(idx, t, T)) \
                                      / self.calculateSigma(idx, t, T)
        return tTest

    def calculateKeppar(self, idx, t, T):
        keppar = self.calculateMiu(idx, t, T) / self.calculateMiu(idx, t-1, T-1) - 1
        return keppar

    def loadDailyReturn(self):
        for idx, item in self.e1_data.iterrows():
            try:
                self.e1_data.loc[idx, 'R'] = self.e1_data.loc[idx, 'Last Trade'] - self.e1_data.loc[idx-1, 'Last Trade']
            except:
                pass

        for idx, item in self.e2_data.iterrows():
            try:
                self.e2_data.loc[idx, 'R'] = self.e2_data.loc[idx, 'Last Trade'] \
                                                     - self.e2_data.loc[idx-1, 'Last Trade']
            except:
                pass

    def keppar(self, term=10, ti=5):
        self.loadDailyReturn()

        for idx, item in self.e1_data.iterrows():
            try:
                self.outdata.loc[idx, 'tTest'] = self.calculateTtest(idx, ti, term)
                self.outdata.loc[idx, 'keppar'] = self.calculateKeppar(idx, ti, term)
            except:
                pass


        for idx, item in self.e1_data.iterrows():
            self.outdata.loc[idx, 'keppar_Enable'] = -0.3 < self.outdata.loc[idx, 'keppar'] and 0.3 > self.outdata.loc[
                idx, 'keppar']
            # LASB = Long A Short B, SALB = Short A Long B
            self.outdata.loc[idx, 'kp_LASB'] = (self.outdata.loc[idx, 'tTest'] > 1.8) * self.outdata.loc[idx, 'keppar_Enable']
            self.outdata.loc[idx, 'kp_SALB'] = (self.outdata.loc[idx, 'tTest'] < 1.8) * self.outdata.loc[idx, 'keppar_Enable']
            self.outdata.loc[idx, 'kp_close'] = ((self.outdata.loc[idx, 'tTest'] < 0 and self.outdata.loc[
                                                     idx - 1, 'tTest'] > 0) or \
                                                 (self.outdata.loc[idx, 'tTest'] > 3 and self.outdata.loc[
                                                     idx - 1, 'tTest'] < 3) or \
                                                 (self.outdata.loc[idx, 'tTest'] < -3 and self.outdata.loc[
                                                     idx - 1, 'tTest'] > -3)) * self.outdata.loc[idx, 'keppar_Enable']

    def draw(self):
        plt.figure(figsize=(12, 6))
        plt.plot(self.e1_data[['Last Trade']])
        plt.plot(self.e2_data[['Last Trade']])
        plt2 = plt.twinx()
        plt2.set_ylim(-10, 10)
        plt2.plot(self.outdata['keppar'], 'y-', markersize=1)
        plt2.plot(self.outdata['tTest'], 'g-', markersize=1)
        plt2.plot(self.outdata['kp_LASB'], 'bo', markersize=1)
        plt2.plot(self.outdata['kp_SALB'], 'ro', markersize=1)
        plt2.plot(self.outdata['kp_close'] * -0.1, 'yo', markersize=1) # will be shown slightly below no-signal
        plt.title('Keppar')
        plt.ylabel('Left - Value; Right - Buy = blue, Sell = red')
        plt.show()

    def getOutData(self):
        self.outdata['E1'] = self.e1_data['Last Trade']
        self.outdata['E2'] = self.e2_data['Last Trade']
        return self.outdata