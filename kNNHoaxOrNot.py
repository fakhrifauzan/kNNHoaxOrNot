import math
import random
import xlrd as x
import xlsxwriter as w

class kNN :
    def __init__(self, training, testing, k):
        self.training = training
        self.testing = testing
        self.hasilHoax = []
        self.hasilSort = []
        self.distance = []
        self.k = k

    def getClasification(self):
        for i in range(len(self.testing)):
            self.hasilSort = []
            self.distance = []
            cYa = 0
            cTidak = 0

            for j in range(len(self.training)):
                d = math.sqrt(((self.training[j][0] - self.testing[i][0]) ** 2) + ((self.training[j][1] - self.testing[i][1]) ** 2) + ((self.training[j][2] - self.testing[i][2]) ** 2) + ((self.training[j][3] - self.testing[i][3]) ** 2))
                self.distance.append([i + 1, j + 1, self.training[j][4], d])

            self.hasilSort = sorted(self.distance, key=lambda distance: distance[3], reverse=True)

            for y in range(self.k):
                if (self.hasilSort[y][2] == 1.0):
                    cYa += 1
                else:
                    cTidak += 1

            if (cYa > cTidak):
                klasifikasi = 1.0
            else:
                klasifikasi = 0.0

            self.hasilHoax.append([self.testing[i][0], self.testing[i][1], self.testing[i][2], self.testing[i][3], klasifikasi])

    def getAccuracy(self):
        cSama = 0.0
        for x in range(len(self.testing)):
            if self.testing[x][4] == self.hasilHoax[x][4]:
                    cSama += 1
        akurasi = (cSama / len(self.testing)) * 100
        return akurasi

    def main(self):
        self.getClasification()
        return self.getAccuracy()

    def printResult(self):
        workbook = w.Workbook('result.xlsx')
        worksheet = workbook.add_worksheet('Result Testing')

        worksheet.write(0, 0, 'Berita')
        worksheet.write(0, 1, 'Like')
        worksheet.write(0, 2, 'Provokasi')
        worksheet.write(0, 3, 'Komentar')
        worksheet.write(0, 4, 'Emosi')
        worksheet.write(0, 5, 'Hoax')

        baris = 1
        kolom = 0
        berita = 4001
        for i in range(len(self.hasilHoax)) :
            worksheet.write(baris, kolom, 'B'+str(berita))
            worksheet.write(baris, kolom+1, self.hasilHoax[i][0])
            worksheet.write(baris, kolom+2, self.hasilHoax[i][1])
            worksheet.write(baris, kolom+3, self.hasilHoax[i][2])
            worksheet.write(baris, kolom+4, self.hasilHoax[i][3])
            worksheet.write(baris, kolom+5, self.hasilHoax[i][4])
            baris += 1
            berita += 1
        print ('data telah disimpan')
        workbook.close()


#Main Program
file = x.open_workbook('data.xlsx')
sheet = file.sheet_by_index(0)
count = sheet.nrows
sheetTest = file.sheet_by_index(1)
countTest = sheetTest.nrows

dataTraining = []
dataTesting = []
akurasi = []

k = 3

for i in range(1,count) :
    dataTraining.append(sheet.row_values(i,1))

for z in range(1,countTest) :
    dataTesting.append(sheetTest.row_values(z,1))

for j in range(k):
    validasi = []
    training = dataTraining
    random.shuffle(training)

    for n in range(1000) :
        validasi.append(training.pop())

    klasifikasi = kNN(training, validasi, k)
    result = klasifikasi.main()
    akurasi.append(result)
    print (result)

print ('Rata-rata : ' + str(sum(akurasi) / float(len(akurasi))))

#cekDataTesting
for x in range(1):
    klasTest = kNN(dataTraining, dataTesting, k)
    klasTest.main()
    klasTest.printResult()