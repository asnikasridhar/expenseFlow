const STORAGE_KEY = "expenseflow_full_github_v1";

const EXCEL_IMPORTED_EXPENSES = [{"id":100021,"date":"2026-05-01","paidBy":"A","amount":1819.0,"category":"Blinkit Wallet","description":"Imported from Excel row 2","payment":"Not specified","reimburse":"No"},{"id":100031,"date":"2026-05-01","paidBy":"A","amount":307.0,"category":"Blinkit Wallet","description":"Imported from Excel row 3","payment":"Not specified","reimburse":"No"},{"id":100041,"date":"2026-05-01","paidBy":"A","amount":305.0,"category":"Compulsory Expenses","description":"Imported from Excel row 4","payment":"Not specified","reimburse":"No"},{"id":100052,"date":"2026-05-01","paidBy":"P","amount":4942.0,"category":"Outing","description":"Imported from Excel row 5","payment":"Not specified","reimburse":"No"},{"id":100062,"date":"2026-05-02","paidBy":"P","amount":256.0,"category":"Kids Supplies","description":"Imported from Excel row 6","payment":"Not specified","reimburse":"No"},{"id":100072,"date":"2026-05-02","paidBy":"P","amount":1350.0,"category":"Kids Supplies","description":"Imported from Excel row 7","payment":"Not specified","reimburse":"No"},{"id":100081,"date":"2026-05-02","paidBy":"A","amount":796.0,"category":"Blinkit Wallet","description":"Imported from Excel row 8","payment":"Not specified","reimburse":"No"},{"id":100091,"date":"2026-05-02","paidBy":"A","amount":940.0,"category":"Outing","description":"Imported from Excel row 9","payment":"Not specified","reimburse":"No"},{"id":100102,"date":"2026-05-02","paidBy":"P","amount":1600.0,"category":"Outing","description":"Imported from Excel row 10","payment":"Not specified","reimburse":"No"},{"id":100112,"date":"2026-05-02","paidBy":"P","amount":5729.0,"category":"Outing","description":"Imported from Excel row 11","payment":"Not specified","reimburse":"No"},{"id":100121,"date":"2026-05-03","paidBy":"A","amount":320.0,"category":"Outing","description":"Imported from Excel row 12","payment":"Not specified","reimburse":"No"},{"id":100131,"date":"2026-05-03","paidBy":"A","amount":1210.0,"category":"Compulsory Expenses","description":"Imported from Excel row 13","payment":"Not specified","reimburse":"No"},{"id":100141,"date":"2026-05-03","paidBy":"A","amount":210.0,"category":"Outing","description":"Imported from Excel row 14","payment":"Not specified","reimburse":"No"},{"id":100151,"date":"2026-05-03","paidBy":"A","amount":638.0,"category":"Blinkit Wallet","description":"Imported from Excel row 15","payment":"Not specified","reimburse":"No"},{"id":100162,"date":"2026-05-03","paidBy":"P","amount":3751.0,"category":"Outing","description":"Imported from Excel row 16","payment":"Not specified","reimburse":"No"},{"id":100172,"date":"2026-05-06","paidBy":"P","amount":1096.0,"category":"Blinkit Wallet","description":"Imported from Excel row 17","payment":"Not specified","reimburse":"No"},{"id":100182,"date":"2026-05-06","paidBy":"P","amount":787.0,"category":"Amazon Wallet","description":"Imported from Excel row 18","payment":"Not specified","reimburse":"No"},{"id":100192,"date":"2026-05-06","paidBy":"P","amount":929.0,"category":"Amazon Wallet","description":"Imported from Excel row 19","payment":"Not specified","reimburse":"No"},{"id":100202,"date":"2026-05-06","paidBy":"P","amount":250.0,"category":"Blinkit Wallet","description":"Imported from Excel row 20","payment":"Not specified","reimburse":"No"},{"id":100212,"date":"2026-05-07","paidBy":"P","amount":6206.0,"category":"Compulsory Expenses","description":"Imported from Excel row 21","payment":"Not specified","reimburse":"No"},{"id":100222,"date":"2026-05-07","paidBy":"P","amount":500.0,"category":"Travel","description":"Imported from Excel row 22","payment":"Not specified","reimburse":"No"},{"id":100231,"date":"2026-05-04","paidBy":"A","amount":1002.0,"category":"Blinkit Wallet","description":"Imported from Excel row 23","payment":"Not specified","reimburse":"No"},{"id":100241,"date":"2026-05-07","paidBy":"A","amount":3910.0,"category":"Blinkit Wallet","description":"Imported from Excel row 24","payment":"Not specified","reimburse":"No"},{"id":100251,"date":"2026-05-07","paidBy":"A","amount":20564.0,"category":"Compulsory Expenses","description":"Imported from Excel row 25","payment":"Not specified","reimburse":"No"},{"id":100261,"date":"2026-05-08","paidBy":"A","amount":622.0,"category":"Blinkit Wallet","description":"Imported from Excel row 26","payment":"Not specified","reimburse":"No"},{"id":100271,"date":"2026-05-05","paidBy":"A","amount":748.0,"category":"Compulsory Expenses","description":"Imported from Excel row 27","payment":"Not specified","reimburse":"No"},{"id":100281,"date":"2026-05-05","paidBy":"A","amount":492.0,"category":"Compulsory Expenses","description":"Imported from Excel row 28","payment":"Not specified","reimburse":"No"},{"id":100291,"date":"2026-05-08","paidBy":"A","amount":745.0,"category":"Compulsory Expenses","description":"Imported from Excel row 29","payment":"Not specified","reimburse":"No"},{"id":100301,"date":"2026-05-08","paidBy":"A","amount":506.0,"category":"Compulsory Expenses","description":"Imported from Excel row 30","payment":"Not specified","reimburse":"No"},{"id":100312,"date":"2026-05-08","paidBy":"P","amount":310.0,"category":"Blinkit Wallet","description":"Imported from Excel row 31","payment":"Not specified","reimburse":"No"},{"id":100322,"date":"2026-05-08","paidBy":"P","amount":837.0,"category":"Amazon Wallet","description":"Imported from Excel row 32","payment":"Not specified","reimburse":"No"},{"id":100331,"date":"2026-05-08","paidBy":"A","amount":10000.0,"category":"Compulsory Expenses","description":"Imported from Excel row 33","payment":"Not specified","reimburse":"No"},{"id":100341,"date":"2026-05-01","paidBy":"A","amount":40000.0,"category":"Home Loan Prepayment","description":"Imported from Excel row 34","payment":"Not specified","reimburse":"No"},{"id":100352,"date":"2026-05-01","paidBy":"P","amount":40000.0,"category":"Home Loan Prepayment","description":"Imported from Excel row 35","payment":"Not specified","reimburse":"No"},{"id":100362,"date":"2026-05-10","paidBy":"P","amount":52700.0,"category":"Home Loan EMI","description":"Imported from Excel row 36","payment":"Not specified","reimburse":"No"},{"id":100372,"date":"2026-05-01","paidBy":"P","amount":35500.0,"category":"Investments (Fixed)","description":"Imported from Excel row 37","payment":"Not specified","reimburse":"No"},{"id":100382,"date":"2026-05-01","paidBy":"P","amount":20000.0,"category":"Investments (Fixed)","description":"Imported from Excel row 38","payment":"Not specified","reimburse":"No"},{"id":100392,"date":"2026-05-08","paidBy":"P","amount":8793.0,"category":"Compulsory Expenses","description":"Imported from Excel row 39","payment":"Not specified","reimburse":"No"},{"id":100402,"date":"2026-05-08","paidBy":"P","amount":199.0,"category":"Compulsory Expenses","description":"Imported from Excel row 40","payment":"Not specified","reimburse":"No"},{"id":100412,"date":"2026-05-01","paidBy":"P","amount":600.0,"category":"Compulsory Expenses","description":"Imported from Excel row 41","payment":"Not specified","reimburse":"No"},{"id":100422,"date":"2026-05-01","paidBy":"P","amount":825.0,"category":"Compulsory Expenses","description":"Imported from Excel row 42","payment":"Not specified","reimburse":"No"},{"id":100431,"date":"2026-05-09","paidBy":"A","amount":805.0,"category":"Blinkit Wallet","description":"Imported from Excel row 43","payment":"Not specified","reimburse":"No"},{"id":100441,"date":"2026-05-09","paidBy":"A","amount":3754.0,"category":"Blinkit Wallet","description":"Imported from Excel row 44","payment":"Not specified","reimburse":"No"},{"id":100452,"date":"2026-05-09","paidBy":"P","amount":4632.0,"category":"Compulsory Expenses","description":"Imported from Excel row 45","payment":"Not specified","reimburse":"No"},{"id":100462,"date":"2026-05-09","paidBy":"P","amount":190.0,"category":"Outing","description":"Imported from Excel row 46","payment":"Not specified","reimburse":"No"},{"id":100472,"date":"2026-05-09","paidBy":"P","amount":329.0,"category":"Outing","description":"Imported from Excel row 47","payment":"Not specified","reimburse":"No"},{"id":100482,"date":"2026-05-09","paidBy":"P","amount":715.0,"category":"Compulsory Expenses","description":"Imported from Excel row 48","payment":"Not specified","reimburse":"No"},{"id":100492,"date":"2026-05-10","paidBy":"P","amount":576.0,"category":"Blinkit Wallet","description":"Imported from Excel row 49","payment":"Not specified","reimburse":"No"},{"id":100502,"date":"2026-05-10","paidBy":"P","amount":3806.0,"category":"Outing","description":"Imported from Excel row 50","payment":"Not specified","reimburse":"No"},{"id":100512,"date":"2026-05-11","paidBy":"P","amount":233.0,"category":"Blinkit Wallet","description":"Imported from Excel row 51","payment":"Not specified","reimburse":"No"},{"id":100522,"date":"2026-05-11","paidBy":"P","amount":5036.0,"category":"Kids Supplies","description":"Imported from Excel row 52","payment":"Not specified","reimburse":"No"},{"id":100531,"date":"2026-05-11","paidBy":"A","amount":4715.0,"category":"Compulsory Expenses","description":"Imported from Excel row 53","payment":"Not specified","reimburse":"No"},{"id":100541,"date":"2026-05-12","paidBy":"A","amount":340.0,"category":"Blinkit Wallet","description":"Imported from Excel row 54","payment":"Not specified","reimburse":"No"},{"id":100542,"date":"2026-05-12","paidBy":"P","amount":2455.0,"category":"Blinkit Wallet","description":"Imported from Excel row 54","payment":"Not specified","reimburse":"No"},{"id":100552,"date":"2026-05-12","paidBy":"P","amount":670.0,"category":"Outing","description":"Imported from Excel row 55","payment":"Not specified","reimburse":"No"},{"id":100561,"date":"2026-05-12","paidBy":"A","amount":474.0,"category":"Blinkit Wallet","description":"Imported from Excel row 56","payment":"Not specified","reimburse":"No"},{"id":100562,"date":"2026-05-12","paidBy":"P","amount":681.0,"category":"Blinkit Wallet","description":"Imported from Excel row 56","payment":"Not specified","reimburse":"No"},{"id":100572,"date":"2026-05-12","paidBy":"P","amount":4012.0,"category":"Travel","description":"Imported from Excel row 57","payment":"Not specified","reimburse":"No"},{"id":100581,"date":"2026-05-13","paidBy":"A","amount":1030.0,"category":"Blinkit Wallet","description":"Imported from Excel row 58","payment":"Not specified","reimburse":"No"},{"id":100582,"date":"2026-05-13","paidBy":"P","amount":200.0,"category":"Blinkit Wallet","description":"Imported from Excel row 58","payment":"Not specified","reimburse":"No"},{"id":100591,"date":"2026-05-13","paidBy":"A","amount":1800.0,"category":"Blinkit Wallet","description":"Imported from Excel row 59","payment":"Not specified","reimburse":"No"},{"id":100592,"date":"2026-05-13","paidBy":"P","amount":500.0,"category":"Blinkit Wallet","description":"Imported from Excel row 59","payment":"Not specified","reimburse":"No"},{"id":100601,"date":"2026-05-13","paidBy":"A","amount":401.0,"category":"Compulsory Expenses","description":"Imported from Excel row 60","payment":"Not specified","reimburse":"No"},{"id":100612,"date":"2026-05-14","paidBy":"P","amount":380.0,"category":"Compulsory Expenses","description":"Imported from Excel row 61","payment":"Not specified","reimburse":"No"},{"id":100622,"date":"2026-05-14","paidBy":"P","amount":456.0,"category":"Compulsory Expenses","description":"Imported from Excel row 62","payment":"Not specified","reimburse":"No"},{"id":100632,"date":"2026-05-14","paidBy":"P","amount":1500.0,"category":"Compulsory Expenses","description":"Imported from Excel row 63","payment":"Not specified","reimburse":"No"},{"id":100642,"date":"2026-05-16","paidBy":"P","amount":250.0,"category":"Compulsory Expenses","description":"Imported from Excel row 64","payment":"Not specified","reimburse":"No"},{"id":100652,"date":"2026-05-17","paidBy":"P","amount":1000.0,"category":"Travel","description":"Imported from Excel row 65","payment":"Not specified","reimburse":"No"},{"id":100662,"date":"2026-05-14","paidBy":"P","amount":951.0,"category":"Outing","description":"Imported from Excel row 66","payment":"Not specified","reimburse":"No"},{"id":100672,"date":"2026-05-15","paidBy":"P","amount":563.0,"category":"Blinkit Wallet","description":"Imported from Excel row 67","payment":"Not specified","reimburse":"No"},{"id":100682,"date":"2026-05-16","paidBy":"P","amount":513.0,"category":"Outing","description":"Imported from Excel row 68","payment":"Not specified","reimburse":"No"},{"id":100692,"date":"2026-05-16","paidBy":"P","amount":928.0,"category":"Outing","description":"Imported from Excel row 69","payment":"Not specified","reimburse":"No"},{"id":100702,"date":"2026-05-17","paidBy":"P","amount":4431.0,"category":"Outing","description":"Imported from Excel row 70","payment":"Not specified","reimburse":"No"},{"id":100712,"date":"2026-05-17","paidBy":"P","amount":285.0,"category":"Blinkit Wallet","description":"Imported from Excel row 71","payment":"Not specified","reimburse":"No"},{"id":100722,"date":"2026-05-09","paidBy":"P","amount":35.0,"category":"Compulsory Expenses","description":"Imported from Excel row 72","payment":"Not specified","reimburse":"No"},{"id":100731,"date":"2026-05-14","paidBy":"A","amount":333.0,"category":"Blinkit Wallet","description":"Imported from Excel row 73","payment":"Not specified","reimburse":"No"},{"id":100741,"date":"2026-05-14","paidBy":"A","amount":744.0,"category":"Outing","description":"Imported from Excel row 74","payment":"Not specified","reimburse":"No"},{"id":100751,"date":"2026-05-14","paidBy":"A","amount":1200.0,"category":"Compulsory Expenses","description":"Imported from Excel row 75","payment":"Not specified","reimburse":"No"},{"id":100761,"date":"2026-05-15","paidBy":"A","amount":377.0,"category":"Blinkit Wallet","description":"Imported from Excel row 76","payment":"Not specified","reimburse":"No"},{"id":100771,"date":"2026-05-16","paidBy":"A","amount":1525.0,"category":"Blinkit Wallet","description":"Imported from Excel row 77","payment":"Not specified","reimburse":"No"},{"id":100781,"date":"2026-05-17","paidBy":"A","amount":936.0,"category":"Outing","description":"Imported from Excel row 78","payment":"Not specified","reimburse":"No"},{"id":100791,"date":"2026-05-19","paidBy":"A","amount":726.0,"category":"Compulsory Expenses","description":"Imported from Excel row 79","payment":"Not specified","reimburse":"No"},{"id":100801,"date":"2026-05-19","paidBy":"A","amount":2414.0,"category":"Blinkit Wallet","description":"Imported from Excel row 80","payment":"Not specified","reimburse":"No"},{"id":100811,"date":"2026-05-20","paidBy":"A","amount":3235.0,"category":"Blinkit Wallet","description":"Imported from Excel row 81","payment":"Not specified","reimburse":"No"},{"id":100821,"date":"2026-05-20","paidBy":"A","amount":676.0,"category":"Compulsory Expenses","description":"Imported from Excel row 82","payment":"Not specified","reimburse":"No"},{"id":100831,"date":"2026-05-21","paidBy":"A","amount":850.0,"category":"Outing","description":"Imported from Excel row 83","payment":"Not specified","reimburse":"No"},{"id":100841,"date":"2026-05-23","paidBy":"A","amount":2514.0,"category":"Outing","description":"Imported from Excel row 84","payment":"Not specified","reimburse":"No"},{"id":100851,"date":"2026-05-23","paidBy":"A","amount":760.0,"category":"Blinkit Wallet","description":"Imported from Excel row 85","payment":"Not specified","reimburse":"No"},{"id":100861,"date":"2026-05-24","paidBy":"A","amount":714.0,"category":"Outing","description":"Imported from Excel row 86","payment":"Not specified","reimburse":"No"},{"id":100871,"date":"2026-05-24","paidBy":"A","amount":1426.0,"category":"Outing","description":"Imported from Excel row 87","payment":"Not specified","reimburse":"No"},{"id":100881,"date":"2026-05-24","paidBy":"A","amount":1178.0,"category":"Compulsory Expenses","description":"Imported from Excel row 88","payment":"Not specified","reimburse":"No"},{"id":100891,"date":"2026-05-25","paidBy":"A","amount":3844.0,"category":"Compulsory Expenses","description":"Imported from Excel row 89","payment":"Not specified","reimburse":"No"},{"id":100901,"date":"2026-05-26","paidBy":"A","amount":325.0,"category":"Blinkit Wallet","description":"Imported from Excel row 90","payment":"Not specified","reimburse":"No"},{"id":100911,"date":"2026-05-28","paidBy":"A","amount":828.0,"category":"Blinkit Wallet","description":"Imported from Excel row 91","payment":"Not specified","reimburse":"No"},{"id":100921,"date":"2026-05-29","paidBy":"A","amount":363.0,"category":"Blinkit Wallet","description":"Imported from Excel row 92","payment":"Not specified","reimburse":"No"},{"id":100931,"date":"2026-05-29","paidBy":"A","amount":16800.0,"category":"Investments (Fixed)","description":"Imported from Excel row 93","payment":"Not specified","reimburse":"No"},{"id":100941,"date":"2026-05-09","paidBy":"A","amount":35.0,"category":"Kids Supplies","description":"Imported from Excel row 94","payment":"Not specified","reimburse":"No"},{"id":100951,"date":"2026-05-09","paidBy":"A","amount":1075.0,"category":"Amazon Wallet","description":"Imported from Excel row 95","payment":"Not specified","reimburse":"No"},{"id":100961,"date":"2026-05-10","paidBy":"A","amount":80.0,"category":"Compulsory Expenses","description":"Imported from Excel row 96","payment":"Not specified","reimburse":"No"},{"id":100971,"date":"2026-05-10","paidBy":"A","amount":2842.0,"category":"Compulsory Expenses","description":"Imported from Excel row 97","payment":"Not specified","reimburse":"No"},{"id":100981,"date":"2026-05-10","paidBy":"A","amount":127.0,"category":"Compulsory Expenses","description":"Imported from Excel row 98","payment":"Not specified","reimburse":"No"},{"id":100991,"date":"2026-05-11","paidBy":"A","amount":747.0,"category":"Compulsory Expenses","description":"Imported from Excel row 99","payment":"Not specified","reimburse":"No"},{"id":101001,"date":"2026-05-11","paidBy":"A","amount":427.0,"category":"Compulsory Expenses","description":"Imported from Excel row 100","payment":"Not specified","reimburse":"No"},{"id":101031,"date":"2026-05-15","paidBy":"A","amount":3800.0,"category":"Compulsory Expenses","description":"Imported from Excel row 103","payment":"Not specified","reimburse":"No"},{"id":101041,"date":"2026-05-16","paidBy":"A","amount":4000.0,"category":"Compulsory Expenses","description":"Imported from Excel row 104","payment":"Not specified","reimburse":"No"},{"id":101051,"date":"2026-05-29","paidBy":"A","amount":540.0,"category":"Blinkit Wallet","description":"Imported from Excel row 105","payment":"Not specified","reimburse":"No"},{"id":101062,"date":"2026-05-17","paidBy":"P","amount":1000.0,"category":"Travel","description":"Imported from Excel row 106","payment":"Not specified","reimburse":"No"},{"id":101072,"date":"2026-05-17","paidBy":"P","amount":95.0,"category":"Compulsory Expenses","description":"Imported from Excel row 107","payment":"Not specified","reimburse":"No"},{"id":101082,"date":"2026-05-18","paidBy":"P","amount":1041.0,"category":"Compulsory Expenses","description":"Imported from Excel row 108","payment":"Not specified","reimburse":"No"},{"id":101092,"date":"2026-05-19","paidBy":"P","amount":300.0,"category":"Compulsory Expenses","description":"Imported from Excel row 109","payment":"Not specified","reimburse":"No"},{"id":101102,"date":"2026-05-20","paidBy":"P","amount":1861.0,"category":"Compulsory Expenses","description":"Imported from Excel row 110","payment":"Not specified","reimburse":"No"},{"id":101112,"date":"2026-05-21","paidBy":"P","amount":1911.0,"category":"Compulsory Expenses","description":"Imported from Excel row 111","payment":"Not specified","reimburse":"No"},{"id":101122,"date":"2026-05-22","paidBy":"P","amount":220.0,"category":"Compulsory Expenses","description":"Imported from Excel row 112","payment":"Not specified","reimburse":"No"},{"id":101132,"date":"2026-05-23","paidBy":"P","amount":650.0,"category":"Travel","description":"Imported from Excel row 113","payment":"Not specified","reimburse":"No"},{"id":101142,"date":"2026-05-23","paidBy":"P","amount":785.0,"category":"Compulsory Expenses","description":"Imported from Excel row 114","payment":"Not specified","reimburse":"No"},{"id":101152,"date":"2026-05-23","paidBy":"P","amount":1055.0,"category":"Outing","description":"Imported from Excel row 115","payment":"Not specified","reimburse":"No"},{"id":101162,"date":"2026-05-24","paidBy":"P","amount":1120.0,"category":"Travel","description":"Imported from Excel row 116","payment":"Not specified","reimburse":"No"},{"id":101172,"date":"2026-05-25","paidBy":"P","amount":500.0,"category":"Kids Supplies","description":"Imported from Excel row 117","payment":"Not specified","reimburse":"No"},{"id":101182,"date":"2026-05-27","paidBy":"P","amount":285.0,"category":"Outing","description":"Imported from Excel row 118","payment":"Not specified","reimburse":"No"},{"id":101192,"date":"2026-05-28","paidBy":"P","amount":815.0,"category":"Compulsory Expenses","description":"Imported from Excel row 119","payment":"Not specified","reimburse":"No"},{"id":101202,"date":"2026-05-29","paidBy":"P","amount":29.0,"category":"Compulsory Expenses","description":"Imported from Excel row 120","payment":"Not specified","reimburse":"No"},{"id":101212,"date":"2026-05-30","paidBy":"P","amount":270.0,"category":"Compulsory Expenses","description":"Imported from Excel row 121","payment":"Not specified","reimburse":"No"},{"id":101222,"date":"2026-05-31","paidBy":"P","amount":275.0,"category":"Compulsory Expenses","description":"Imported from Excel row 122","payment":"Not specified","reimburse":"No"},{"id":101232,"date":"2026-05-26","paidBy":"P","amount":1608.0,"category":"Compulsory Expenses","description":"Imported from Excel row 123","payment":"Not specified","reimburse":"No"},{"id":101242,"date":"2026-05-27","paidBy":"P","amount":732.0,"category":"Outing","description":"Imported from Excel row 124","payment":"Not specified","reimburse":"No"},{"id":101252,"date":"2026-05-28","paidBy":"P","amount":279.0,"category":"Compulsory Expenses","description":"Imported from Excel row 125","payment":"Not specified","reimburse":"No"},{"id":101262,"date":"2026-05-29","paidBy":"P","amount":1105.0,"category":"Compulsory Expenses","description":"Imported from Excel row 126","payment":"Not specified","reimburse":"No"},{"id":101272,"date":"2026-05-30","paidBy":"P","amount":1214.0,"category":"Blinkit Wallet","description":"Imported from Excel row 127","payment":"Not specified","reimburse":"No"},{"id":101282,"date":"2026-05-31","paidBy":"P","amount":1023.0,"category":"Blinkit Wallet","description":"Imported from Excel row 128","payment":"Not specified","reimburse":"No"},{"id":101292,"date":"2026-05-31","paidBy":"P","amount":825.0,"category":"Compulsory Expenses","description":"Imported from Excel row 129","payment":"Not specified","reimburse":"No"},{"id":101302,"date":"2026-05-18","paidBy":"P","amount":513.0,"category":"Blinkit Wallet","description":"Imported from Excel row 130","payment":"Not specified","reimburse":"No"},{"id":101312,"date":"2026-05-19","paidBy":"P","amount":669.0,"category":"Blinkit Wallet","description":"Imported from Excel row 131","payment":"Not specified","reimburse":"No"},{"id":101322,"date":"2026-05-24","paidBy":"P","amount":110.0,"category":"Compulsory Expenses","description":"Imported from Excel row 132","payment":"Not specified","reimburse":"No"},{"id":101332,"date":"2026-05-24","paidBy":"P","amount":4322.0,"category":"Outing","description":"Imported from Excel row 133","payment":"Not specified","reimburse":"No"},{"id":101342,"date":"2026-05-25","paidBy":"P","amount":3445.0,"category":"Travel","description":"Imported from Excel row 134","payment":"Not specified","reimburse":"No"},{"id":101352,"date":"2026-05-21","paidBy":"P","amount":17600.0,"category":"Travel","description":"Imported from Excel row 135","payment":"Not specified","reimburse":"No"},{"id":101371,"date":"2026-06-01","paidBy":"A","amount":10500.0,"category":"Compulsory Expenses","description":"Imported from Excel row 137","payment":"Not specified","reimburse":"No"},{"id":101381,"date":"2026-06-02","paidBy":"A","amount":6849.0,"category":"Compulsory Expenses","description":"Imported from Excel row 138","payment":"Not specified","reimburse":"No"},{"id":101391,"date":"2026-06-02","paidBy":"A","amount":20000.0,"category":"Investments (Fixed)","description":"Imported from Excel row 139","payment":"Not specified","reimburse":"No"},{"id":101401,"date":"2026-06-03","paidBy":"A","amount":650.0,"category":"Compulsory Expenses","description":"Imported from Excel row 140","payment":"Not specified","reimburse":"No"},{"id":101411,"date":"2026-06-02","paidBy":"A","amount":40000.0,"category":"Home Loan Prepayment","description":"Imported from Excel row 141","payment":"Not specified","reimburse":"No"},{"id":101421,"date":"2026-06-02","paidBy":"A","amount":858.0,"category":"Outing","description":"Imported from Excel row 142","payment":"Not specified","reimburse":"No"},{"id":101431,"date":"2026-06-03","paidBy":"A","amount":10000.0,"category":"Compulsory Expenses","description":"Imported from Excel row 143","payment":"Not specified","reimburse":"No"},{"id":101441,"date":"2026-06-02","paidBy":"A","amount":675.0,"category":"Blinkit Wallet","description":"Imported from Excel row 144","payment":"Not specified","reimburse":"No"},{"id":101451,"date":"2026-06-03","paidBy":"A","amount":582.0,"category":"Compulsory Expenses","description":"Imported from Excel row 145","payment":"Not specified","reimburse":"No"},{"id":101461,"date":"2026-06-02","paidBy":"A","amount":178.0,"category":"Blinkit Wallet","description":"Imported from Excel row 146","payment":"Not specified","reimburse":"No"},{"id":101471,"date":"2026-06-03","paidBy":"A","amount":460.0,"category":"Compulsory Expenses","description":"Imported from Excel row 147","payment":"Not specified","reimburse":"No"},{"id":101481,"date":"2026-06-02","paidBy":"A","amount":4385.0,"category":"Compulsory Expenses","description":"Imported from Excel row 148","payment":"Not specified","reimburse":"No"},{"id":101491,"date":"2026-06-03","paidBy":"A","amount":60.0,"category":"Compulsory Expenses","description":"Imported from Excel row 149","payment":"Not specified","reimburse":"No"},{"id":101501,"date":"2026-06-03","paidBy":"A","amount":120.0,"category":"Compulsory Expenses","description":"Imported from Excel row 150","payment":"Not specified","reimburse":"No"},{"id":101511,"date":"2026-06-03","paidBy":"A","amount":680.0,"category":"Compulsory Expenses","description":"Imported from Excel row 151","payment":"Not specified","reimburse":"No"},{"id":101522,"date":"2026-06-01","paidBy":"P","amount":2856.0,"category":"Compulsory Expenses","description":"Imported from Excel row 152","payment":"Not specified","reimburse":"No"},{"id":101532,"date":"2026-06-02","paidBy":"P","amount":40000.0,"category":"Home Loan Prepayment","description":"Imported from Excel row 153","payment":"Not specified","reimburse":"No"},{"id":101542,"date":"2026-06-02","paidBy":"P","amount":600.0,"category":"Compulsory Expenses","description":"Imported from Excel row 154","payment":"Not specified","reimburse":"No"},{"id":101552,"date":"2026-06-03","paidBy":"P","amount":2045.0,"category":"Compulsory Expenses","description":"Imported from Excel row 155","payment":"Not specified","reimburse":"No"},{"id":101562,"date":"2026-06-03","paidBy":"P","amount":1272.0,"category":"Blinkit Wallet","description":"Imported from Excel row 156","payment":"Not specified","reimburse":"No"},{"id":101572,"date":"2026-06-03","paidBy":"P","amount":673.0,"category":"Outing","description":"Imported from Excel row 157","payment":"Not specified","reimburse":"No"}];

const defaultState = {
  activeMonth: "2026-05",
  people: { a: "A", p: "P", incomeA: 140000, incomeP: 223000 },
  loan: { original: 3300000, balance: 3150000, emi: 52700, rate: 7.5 },
  budgets: {
    "Home Loan EMI": 52700,
    "Home Loan Prepayment": 80000,
    "Compulsory Expenses": 100000,
    "Kids Supplies": 25000,
    "Travel": 10000,
    "Blinkit Wallet": 16000,
    "Amazon Wallet": 16000,
    "Outing": 8000,
    "Investments (Fixed)": 55000
  },
  prepayments: [
    { id: 1, name: "Regular Prepayment", amount: 80000, frequency: "Monthly" },
    { id: 2, name: "Bonus Prepayment", amount: 100000, frequency: "Occasional" }
  ],
  expenses: []
};

let state = loadState();

function loadState(){
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY)) || structuredClone(defaultState);
  } catch {
    return structuredClone(defaultState);
  }
}

function saveState(){
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

const el = id => document.getElementById(id);
const rupee = n => "₹" + Number(n || 0).toLocaleString("en-IN", { maximumFractionDigits: 0 });
const todayISO = () => new Date().toISOString().slice(0,10);
const monthOf = date => (date || todayISO()).slice(0,7);
const monthName = ym => new Date(ym + "-01T00:00:00").toLocaleString("en-IN", { month:"long", year:"numeric" });

function expensesForMonth(ym = state.activeMonth){
  return state.expenses.filter(e => monthOf(e.date) === ym);
}

function previousMonth(ym){
  const d = new Date(ym + "-01T00:00:00");
  d.setMonth(d.getMonth() - 1);
  return d.toISOString().slice(0,7);
}

function nextMonthValue(ym){
  const d = new Date(ym + "-01T00:00:00");
  d.setMonth(d.getMonth() + 1);
  return d.toISOString().slice(0,7);
}

function totalsByCategory(list){
  return list.reduce((acc,e) => {
    acc[e.category] = (acc[e.category] || 0) + Number(e.amount || 0);
    return acc;
  }, {});
}

function totalsByPerson(list){
  return list.reduce((acc,e) => {
    acc[e.paidBy] = (acc[e.paidBy] || 0) + Number(e.amount || 0);
    return acc;
  }, { A:0, P:0 });
}

function monthSummary(ym){
  const list = expensesForMonth(ym);
  const income = Number(state.people.incomeA || 0) + Number(state.people.incomeP || 0);
  const expense = list.reduce((s,e) => s + Number(e.amount || 0), 0);
  const totalBudget = Object.values(state.budgets).reduce((s,v) => s + Number(v || 0), 0);
  return {
    ym, list, income, expense, surplus: income - expense,
    savingsRate: income ? ((income - expense) / income) * 100 : 0,
    budgetUsed: totalBudget ? (expense / totalBudget) * 100 : 0,
    totalBudget
  };
}


function ensureImportedExcelData(){
  const existingIds = new Set((state.expenses || []).map(e => Number(e.id)));
  const missing = EXCEL_IMPORTED_EXPENSES.filter(e => !existingIds.has(Number(e.id)));
  if (missing.length) {
    state.expenses = [...(state.expenses || []), ...missing].sort((a,b) => String(b.date).localeCompare(String(a.date)));
    if (!state.activeMonth) state.activeMonth = "2026-05";
    saveState();
  }
}

function init(){
  ensureImportedExcelData();
  el("activeMonth").value = state.activeMonth;
  el("date").value = todayISO();
  wireEvents();
  refreshDropdowns();
  renderAll();
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(()=>{});
}

function wireEvents(){
  document.querySelectorAll(".nav-btn").forEach(btn => btn.addEventListener("click", () => showScreen(btn.dataset.screen)));
  el("activeMonth").addEventListener("change", e => { state.activeMonth = e.target.value; saveState(); renderAll(); });
  el("prevMonth").onclick = () => changeMonth(previousMonth(state.activeMonth));
  el("nextMonth").onclick = () => changeMonth(nextMonthValue(state.activeMonth));
  el("expenseForm").addEventListener("submit", saveExpense);
  el("cancelEditBtn").onclick = cancelExpenseEdit;
  el("expenseFilterCategory").onchange = renderExpenseList;
  el("expenseFilterPerson").onchange = renderExpenseList;
  el("expenseSearch").oninput = renderExpenseList;
  el("clearMonthBtn").onclick = clearSelectedMonth;
  el("addBudgetItemBtn").onclick = addBudgetItem;
  el("savePartnerBtn").onclick = savePartner;
  el("saveLoanBtn").onclick = saveLoan;
  el("addPrepayBtn").onclick = addPrepay;
  el("printReportBtn").onclick = () => window.print();
  el("printReportTop").onclick = () => { showScreen("reports"); setTimeout(() => window.print(), 100); };
  el("exportCsvBtn").onclick = exportCSV;
  el("backupJsonBtn").onclick = backupJSON;
  el("restoreJsonInput").onchange = restoreJSON;
  el("seedDemoBtn").onclick = seedDemoData;
  el("resetAllBtn").onclick = resetAll;
}

function changeMonth(ym){
  state.activeMonth = ym;
  el("activeMonth").value = ym;
  saveState();
  renderAll();
}

function showScreen(screenId){
  document.querySelectorAll(".screen").forEach(s => s.classList.remove("active"));
  document.querySelectorAll(".nav-btn").forEach(b => b.classList.toggle("active", b.dataset.screen === screenId));
  el(screenId).classList.add("active");
  renderAll();
}

function refreshDropdowns(){
  el("paidBy").innerHTML = `<option value="A">${state.people.a}</option><option value="P">${state.people.p}</option>`;
  el("category").innerHTML = Object.keys(state.budgets).map(c => `<option>${escapeHtml(c)}</option>`).join("");
  el("expenseFilterCategory").innerHTML = `<option value="">All categories</option>` + Object.keys(state.budgets).map(c => `<option>${escapeHtml(c)}</option>`).join("");
  el("expenseFilterPerson").innerHTML = `<option value="">All people</option><option value="A">${state.people.a}</option><option value="P">${state.people.p}</option>`;
}

function renderAll(){
  refreshDropdowns();
  renderDashboard();
  renderQuickAdd();
  renderExpenseList();
  renderBudgetEditor();
  renderPartner();
  renderLoan();
  renderReport();
}

function renderDashboard(){
  const s = monthSummary(state.activeMonth);
  el("monthLabel").textContent = monthName(state.activeMonth);
  el("heroSurplus").textContent = rupee(s.surplus);
  el("dashIncome").textContent = rupee(s.income);
  el("dashExpense").textContent = rupee(s.expense);
  el("dashRate").textContent = Math.round(s.savingsRate) + "%";
  el("dashBudgetUsed").textContent = Math.round(s.budgetUsed) + "%";

  renderBudgetBars(s.list);
  renderPartnerBars(s.list);
  renderSurplusBox(s);
  renderLoanSnapshot();
  renderMonthCompare();
}

function renderBudgetBars(list){
  const totals = totalsByCategory(list);
  el("budgetBars").innerHTML = Object.keys(state.budgets).map(cat => {
    const spent = totals[cat] || 0;
    const budget = Number(state.budgets[cat] || 0);
    const pct = budget ? Math.min(100, spent / budget * 100) : 0;
    return `<div class="bar-row ${spent > budget ? "over":""}">
      <div class="bar-top"><span>${escapeHtml(cat)}</span><span>${rupee(spent)} / ${rupee(budget)}</span></div>
      <div class="bar"><i style="width:${pct}%"></i></div>
    </div>`;
  }).join("");
}

function renderPartnerBars(list){
  const t = totalsByPerson(list);
  const max = Math.max(t.A, t.P, 1);
  el("partnerBars").innerHTML = `
    <div class="bar-row"><div class="bar-top"><span>${escapeHtml(state.people.a)}</span><span>${rupee(t.A)}</span></div><div class="bar"><i style="width:${t.A/max*100}%"></i></div></div>
    <div class="bar-row"><div class="bar-top"><span>${escapeHtml(state.people.p)}</span><span>${rupee(t.P)}</span></div><div class="bar"><i style="width:${t.P/max*100}%"></i></div></div>
  `;
}

function renderSurplusBox(s){
  el("surplusBox").innerHTML = `
    <table class="table">
      <tr><td>Monthly surplus</td><td><b>${rupee(s.surplus)}</b></td></tr>
      <tr><td>Liquidity 40%</td><td>${rupee(s.surplus * 0.4)}</td></tr>
      <tr><td>Investment Boost 60%</td><td>${rupee(s.surplus * 0.6)}</td></tr>
    </table>
  `;
}

function renderLoanSnapshot(){
  const paidPct = state.loan.original ? Math.max(0, Math.min(100, (state.loan.original - state.loan.balance) / state.loan.original * 100)) : 0;
  const prepayTotal = state.prepayments.reduce((s,p) => s + Number(p.amount || 0), 0);
  el("loanSnapshot").innerHTML = `
    <div class="bar-row">
      <div class="bar-top"><span>Paid progress</span><span>${Math.round(paidPct)}%</span></div>
      <div class="bar"><i style="width:${paidPct}%"></i></div>
    </div>
    <table class="table">
      <tr><td>Outstanding</td><td>${rupee(state.loan.balance)}</td></tr>
      <tr><td>EMI</td><td>${rupee(state.loan.emi)}</td></tr>
      <tr><td>Prepayment options total</td><td>${rupee(prepayTotal)}</td></tr>
      <tr><td>Interest rate</td><td>${state.loan.rate}%</td></tr>
    </table>
  `;
}

function renderMonthCompare(){
  const curr = monthSummary(state.activeMonth);
  const prev = monthSummary(previousMonth(state.activeMonth));
  el("monthCompare").innerHTML = `
    <table class="table">
      <thead><tr><th>Metric</th><th>${monthName(prev.ym)}</th><th>${monthName(curr.ym)}</th></tr></thead>
      <tbody>
        <tr><td>Income</td><td>${rupee(prev.income)}</td><td>${rupee(curr.income)}</td></tr>
        <tr><td>Expense</td><td>${rupee(prev.expense)}</td><td>${rupee(curr.expense)}</td></tr>
        <tr><td>Surplus</td><td>${rupee(prev.surplus)}</td><td>${rupee(curr.surplus)}</td></tr>
        <tr><td>Savings Rate</td><td>${Math.round(prev.savingsRate)}%</td><td>${Math.round(curr.savingsRate)}%</td></tr>
      </tbody>
    </table>
  `;
}

function renderQuickAdd(){
  const quickItems = [
    ...state.prepayments.map(p => ({ label:p.name, amount:p.amount, category: state.budgets[p.name] !== undefined ? p.name : "Home Loan Prepayment", desc:p.name })),
    { label:"EMI", amount:state.loan.emi, category:"Home Loan EMI", desc:"Monthly EMI" },
    { label:"Groceries", amount:1000, category:"Groceries", desc:"Groceries" },
    { label:"Blinkit", amount:1000, category:"Blinkit Wallet", desc:"Blinkit Wallet" },
    { label:"Amazon", amount:1000, category:"Amazon Wallet", desc:"Amazon Wallet" }
  ];
  el("quickAddList").innerHTML = quickItems.map((q, idx) => `
    <button class="quick-card" onclick="fillQuickAdd(${idx})">
      <strong>${escapeHtml(q.label)}</strong>
      <small>${rupee(q.amount)} • ${escapeHtml(q.category)}</small>
    </button>
  `).join("");
  window.__quickItems = quickItems;
}

function fillQuickAdd(idx){
  const q = window.__quickItems[idx];
  showScreen("add");
  el("amount").value = q.amount;
  el("category").value = q.category;
  el("description").value = q.desc;
  el("payment").value = "Bank Transfer";
  el("date").value = state.activeMonth + "-01";
  window.scrollTo(0,0);
}

function saveExpense(e){
  e.preventDefault();
  const id = el("editingExpenseId").value ? Number(el("editingExpenseId").value) : Date.now();
  const expense = {
    id,
    paidBy: el("paidBy").value,
    amount: Number(el("amount").value || 0),
    category: el("category").value,
    payment: el("payment").value,
    date: el("date").value,
    reimburse: el("reimburse").value,
    description: el("description").value.trim()
  };
  const idx = state.expenses.findIndex(x => x.id === id);
  if (idx >= 0) state.expenses[idx] = expense;
  else state.expenses.unshift(expense);
  saveState();
  cancelExpenseEdit();
  state.activeMonth = monthOf(expense.date);
  ensureImportedExcelData();
  el("activeMonth").value = state.activeMonth;
  renderAll();
  showScreen("expenses");
}

function editExpense(id){
  const e = state.expenses.find(x => x.id === id);
  if (!e) return;
  showScreen("add");
  el("editingExpenseId").value = e.id;
  el("paidBy").value = e.paidBy;
  el("amount").value = e.amount;
  el("category").value = e.category;
  el("payment").value = e.payment;
  el("date").value = e.date;
  el("reimburse").value = e.reimburse;
  el("description").value = e.description || "";
  el("expenseFormTitle").textContent = "Edit Expense";
  el("saveExpenseBtn").textContent = "Update Expense";
  el("cancelEditBtn").classList.remove("hidden");
  window.scrollTo(0,0);
}

function deleteExpense(id){
  if (!confirm("Delete this expense?")) return;
  state.expenses = state.expenses.filter(e => e.id !== id);
  saveState();
  renderAll();
}

function cancelExpenseEdit(){
  el("expenseForm").reset();
  el("editingExpenseId").value = "";
  el("date").value = todayISO();
  el("expenseFormTitle").textContent = "Add Daily Expense";
  el("saveExpenseBtn").textContent = "Save Expense";
  el("cancelEditBtn").classList.add("hidden");
  refreshDropdowns();
}

function renderExpenseList(){
  const catFilter = el("expenseFilterCategory").value;
  const personFilter = el("expenseFilterPerson").value;
  const search = el("expenseSearch").value.toLowerCase();
  let list = expensesForMonth();
  if (catFilter) list = list.filter(e => e.category === catFilter);
  if (personFilter) list = list.filter(e => e.paidBy === personFilter);
  if (search) list = list.filter(e => [e.category,e.description,e.payment].join(" ").toLowerCase().includes(search));

  el("expenseList").innerHTML = list.map(e => `
    <div class="entry">
      <div>
        <strong>${escapeHtml(e.category)}</strong><br/>
        <small>${escapeHtml(e.description || e.payment)} • ${e.date}</small><br/>
        <span class="badge">Paid by ${e.paidBy === "A" ? escapeHtml(state.people.a) : escapeHtml(state.people.p)}${e.reimburse === "Yes" ? " • Reimburse" : ""}</span>
        <div class="entry-actions">
          <button class="secondary" onclick="editExpense(${e.id})">Edit</button>
          <button class="danger" onclick="deleteExpense(${e.id})">Delete</button>
        </div>
      </div>
      <div class="amt">${rupee(e.amount)}</div>
    </div>
  `).join("") || `<p>No expenses for ${monthName(state.activeMonth)}.</p>`;
}

function clearSelectedMonth(){
  if (!confirm(`Clear all expenses for ${monthName(state.activeMonth)}?`)) return;
  state.expenses = state.expenses.filter(e => monthOf(e.date) !== state.activeMonth);
  saveState();
  renderAll();
}

function renderBudgetEditor(){
  const list = expensesForMonth();
  const totals = totalsByCategory(list);
  el("budgetEditor").innerHTML = Object.keys(state.budgets).map(cat => {
    const spent = totals[cat] || 0;
    const budget = state.budgets[cat] || 0;
    const pct = budget ? Math.min(100, spent / budget * 100) : 0;
    return `<div class="bar-row ${spent > budget ? "over":""}">
      <div class="bar-top">
        <span>${escapeHtml(cat)}</span>
        <span>${rupee(spent)} / <input class="budget-input" type="number" value="${budget}" onchange="updateBudget('${jsString(cat)}', this.value)"></span>
      </div>
      <div class="bar"><i style="width:${pct}%"></i></div>
      <div class="entry-actions">
        <button class="secondary" onclick="renameBudgetItem('${jsString(cat)}')">Rename</button>
        <button class="danger" onclick="deleteBudgetItem('${jsString(cat)}')">Delete</button>
      </div>
    </div>`;
  }).join("");
}

function updateBudget(cat, value){
  state.budgets[cat] = Number(value || 0);
  saveState();
  renderAll();
}

function addBudgetItem(){
  const name = prompt("Budget item name");
  if (!name) return;
  const amount = Number(prompt("Monthly budget amount", "0") || 0);
  state.budgets[name] = amount;
  saveState();
  renderAll();
}

function renameBudgetItem(oldName){
  const newName = prompt("Rename budget item", oldName);
  if (!newName || newName === oldName) return;
  state.budgets[newName] = state.budgets[oldName];
  delete state.budgets[oldName];
  state.expenses.forEach(e => { if (e.category === oldName) e.category = newName; });
  state.prepayments.forEach(p => { if (p.name === oldName) p.name = newName; });
  saveState();
  renderAll();
}

function deleteBudgetItem(name){
  if (!confirm("Delete this budget item? Existing old expenses will remain.")) return;
  delete state.budgets[name];
  saveState();
  renderAll();
}

function renderPartner(){
  el("personAName").value = state.people.a;
  el("personPName").value = state.people.p;
  el("incomeA").value = state.people.incomeA;
  el("incomeP").value = state.people.incomeP;

  const t = totalsByPerson(expensesForMonth());
  const total = t.A + t.P;
  el("partnerSummary").innerHTML = `
    <table class="table">
      <tr><td>${escapeHtml(state.people.a)} Income</td><td>${rupee(state.people.incomeA)}</td></tr>
      <tr><td>${escapeHtml(state.people.p)} Income</td><td>${rupee(state.people.incomeP)}</td></tr>
      <tr><td>${escapeHtml(state.people.a)} Expenses</td><td>${rupee(t.A)}</td></tr>
      <tr><td>${escapeHtml(state.people.p)} Expenses</td><td>${rupee(t.P)}</td></tr>
      <tr><td>${escapeHtml(state.people.a)} Expense Share</td><td>${total ? Math.round(t.A/total*100) : 0}%</td></tr>
      <tr><td>${escapeHtml(state.people.p)} Expense Share</td><td>${total ? Math.round(t.P/total*100) : 0}%</td></tr>
    </table>
  `;
}

function savePartner(){
  state.people.a = el("personAName").value || "A";
  state.people.p = el("personPName").value || "P";
  state.people.incomeA = Number(el("incomeA").value || 0);
  state.people.incomeP = Number(el("incomeP").value || 0);
  saveState();
  renderAll();
  alert("Partner details saved");
}

function renderLoan(){
  el("loanOriginal").value = state.loan.original;
  el("loanBalance").value = state.loan.balance;
  el("loanEmi").value = state.loan.emi;
  el("loanRate").value = state.loan.rate;

  el("prepayList").innerHTML = state.prepayments.map((p,i) => `
    <div class="entry">
      <div>
        <strong>${escapeHtml(p.name)}</strong><br/>
        <small>${rupee(p.amount)} • ${escapeHtml(p.frequency)}</small>
      </div>
      <div class="entry-actions">
        <button class="secondary" onclick="editPrepay(${i})">Edit</button>
        <button class="secondary" onclick="addPrepayAsExpense(${i})">Add Expense</button>
        <button class="danger" onclick="deletePrepay(${i})">Delete</button>
      </div>
    </div>
  `).join("") || "<p>No prepayment options yet.</p>";
}

function saveLoan(){
  state.loan.original = Number(el("loanOriginal").value || 0);
  state.loan.balance = Number(el("loanBalance").value || 0);
  state.loan.emi = Number(el("loanEmi").value || 0);
  state.loan.rate = Number(el("loanRate").value || 0);
  saveState();
  renderAll();
  alert("Loan saved");
}

function addPrepay(){
  const name = prompt("Prepayment name", "Extra Prepayment");
  if (!name) return;
  const amount = Number(prompt("Amount", "50000") || 0);
  const frequency = prompt("Frequency", "Monthly") || "Monthly";
  state.prepayments.push({ id: Date.now(), name, amount, frequency });
  if (state.budgets[name] === undefined) state.budgets[name] = amount;
  saveState();
  renderAll();
}

function editPrepay(i){
  const p = state.prepayments[i];
  const name = prompt("Prepayment name", p.name);
  if (!name) return;
  const amount = Number(prompt("Amount", String(p.amount)) || 0);
  const frequency = prompt("Frequency", p.frequency) || p.frequency;
  state.prepayments[i] = { ...p, name, amount, frequency };
  if (state.budgets[name] === undefined) state.budgets[name] = amount;
  saveState();
  renderAll();
}

function deletePrepay(i){
  if (!confirm("Delete this prepayment option?")) return;
  state.prepayments.splice(i,1);
  saveState();
  renderAll();
}

function addPrepayAsExpense(i){
  const p = state.prepayments[i];
  showScreen("add");
  el("amount").value = p.amount;
  el("category").value = state.budgets[p.name] !== undefined ? p.name : "Home Loan Prepayment";
  el("description").value = p.name;
  el("payment").value = "Bank Transfer";
  el("date").value = state.activeMonth + "-01";
  window.scrollTo(0,0);
}

function renderReport(){
  const s = monthSummary(state.activeMonth);
  const catTotals = totalsByCategory(s.list);
  const personTotals = totalsByPerson(s.list);
  const topCats = Object.entries(catTotals).sort((a,b) => b[1]-a[1]).slice(0,8);
  el("reportTitle").textContent = `Monthly Finance Report - ${monthName(state.activeMonth)}`;
  el("reportContent").innerHTML = `
    <h3>1. Income Summary</h3>
    <table class="table">
      <tr><td>${escapeHtml(state.people.a)}</td><td>${rupee(state.people.incomeA)}</td></tr>
      <tr><td>${escapeHtml(state.people.p)}</td><td>${rupee(state.people.incomeP)}</td></tr>
      <tr><td><b>Total Income</b></td><td><b>${rupee(s.income)}</b></td></tr>
    </table>

    <h3>2. Expense Summary</h3>
    <table class="table">
      <tr><td>Total Expenses</td><td>${rupee(s.expense)}</td></tr>
      <tr><td>Surplus</td><td>${rupee(s.surplus)}</td></tr>
      <tr><td>Savings Rate</td><td>${Math.round(s.savingsRate)}%</td></tr>
      <tr><td>Budget Used</td><td>${Math.round(s.budgetUsed)}%</td></tr>
    </table>

    <h3>3. Top Categories</h3>
    <table class="table">
      ${topCats.map(([c,v]) => `<tr><td>${escapeHtml(c)}</td><td>${rupee(v)}</td></tr>`).join("") || "<tr><td>No data</td><td></td></tr>"}
    </table>

    <h3>4. Partner Contribution</h3>
    <table class="table">
      <tr><td>${escapeHtml(state.people.a)} paid</td><td>${rupee(personTotals.A)}</td></tr>
      <tr><td>${escapeHtml(state.people.p)} paid</td><td>${rupee(personTotals.P)}</td></tr>
    </table>

    <h3>5. Loan Progress</h3>
    <table class="table">
      <tr><td>Outstanding</td><td>${rupee(state.loan.balance)}</td></tr>
      <tr><td>EMI</td><td>${rupee(state.loan.emi)}</td></tr>
      <tr><td>Interest Rate</td><td>${state.loan.rate}%</td></tr>
    </table>

    <h3>6. Surplus Allocation</h3>
    <table class="table">
      <tr><td>Liquidity 40%</td><td>${rupee(s.surplus * 0.4)}</td></tr>
      <tr><td>Investment Boost 60%</td><td>${rupee(s.surplus * 0.6)}</td></tr>
    </table>
  `;
}

function exportCSV(){
  const rows = [["Date","Month","Paid By","Amount","Category","Description","Payment","Reimburse"]];
  state.expenses.forEach(e => rows.push([
    e.date, monthOf(e.date), e.paidBy === "A" ? state.people.a : state.people.p, e.amount, e.category, e.description, e.payment, e.reimburse
  ]));
  download("expenseflow-expenses.csv", rows.map(r => r.map(v => `"${String(v ?? "").replaceAll('"','""')}"`).join(",")).join("\n"), "text/csv");
}

function backupJSON(){
  download("expenseflow-backup.json", JSON.stringify(state,null,2), "application/json");
}

function restoreJSON(e){
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      state = JSON.parse(reader.result);
      saveState();
      el("activeMonth").value = state.activeMonth || todayISO().slice(0,7);
      renderAll();
      alert("Backup restored");
    } catch {
      alert("Invalid JSON backup");
    }
  };
  reader.readAsText(file);
}

function seedDemoData(){
  if (!confirm("Add demo data for current and previous month?")) return;
  const m = state.activeMonth;
  const pm = previousMonth(m);
  const demo = [
    [m+"-02","A",12000,"Groceries","Monthly groceries","UPI"],
    [m+"-03","P",52700,"Home Loan EMI","Monthly EMI","Bank Transfer"],
    [m+"-05","A",80000,"Home Loan Prepayment","Regular prepayment","Bank Transfer"],
    [m+"-08","P",6000,"Kids Supplies","School and child supplies","UPI"],
    [m+"-11","A",3500,"Outing","Family dinner","Card"],
    [pm+"-02","A",14000,"Groceries","Monthly groceries","UPI"],
    [pm+"-03","P",52700,"Home Loan EMI","Monthly EMI","Bank Transfer"],
    [pm+"-05","A",50000,"Home Loan Prepayment","Prepayment","Bank Transfer"],
    [pm+"-12","P",8000,"Travel","Fuel and travel","UPI"]
  ];
  demo.forEach(d => state.expenses.unshift({ id: Date.now()+Math.random(), date:d[0], paidBy:d[1], amount:d[2], category:d[3], description:d[4], payment:d[5], reimburse:"No" }));
  saveState();
  renderAll();
}

function resetAll(){
  if (!confirm("Reset the full app data on this browser?")) return;
  state = structuredClone(defaultState);
  saveState();
  ensureImportedExcelData();
  el("activeMonth").value = state.activeMonth;
  renderAll();
}

function download(name, content, type){
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([content], { type }));
  a.download = name;
  a.click();
}

function escapeHtml(str){
  return String(str ?? "").replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[s]));
}

function jsString(str){
  return String(str).replace(/\\/g,"\\\\").replace(/'/g,"\\'");
}

init();
