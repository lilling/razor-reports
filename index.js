const axios = require('axios');
const Excel = require('exceljs/modern.nodejs');
const fs = require('fs');

if (!process.argv[2]) {
    console.error('no .xlsx file to load');
    return;
}

const dateParser = function (day, time = null) {
    let minutes = '00';
    let hour = '00';
    if (time) {
        hour = time.getUTCHours().toString();
        if (hour.length === 1) {
            hour = '0' + hour;
        }
        minutes = time.getUTCMinutes().toString();
        if (minutes.length === 1) {
            minutes = '0' + minutes;
        }
    }

    return day + 'T' + hour + ':' + minutes;
};

const bootstrap = function () {
    const url = 'https://hour.razorgrip.com/';
    fs.readFile('user-data', (err, data) => {
        if(err) {
            console.error('make sure you have user-data file')
            return;
        }
		
        const cred = data.toString().split('\r\n').reduce((result, current) => {
            const prop = current.split(':');
            if (prop[1]) {
                result[prop[0]] = prop[1];
            }
            return result;
        }, {});

        axios.post(url + 'login', cred)
            .then(res => {
                const { token, userId } = res.data;
                const vacation = url + 'hourreports/saveVacation';
                const regular = url + 'hourreports/save';
                axios.defaults.headers.common = {'Authorization': `Bearer ${token}`};

                const workbook = new Excel.Workbook();
                workbook.xlsx.readFile(process.argv[2])
                    .then(() => {
                        const worksheet = workbook.getWorksheet(1);
                        worksheet.eachRow((row, rowNumber) => {
                            if (rowNumber < 4) return;
                            const day = new Date().getFullYear() + '-' + row.values[1].split(' ')[1].split('/').reverse().join('-');
                            let data = { user_id: userId, project_id: 40 };

                            if (row.values[9] === 'חופשה') {
                                data.date = dateParser(day);
                            } else {
                                data.from_date = dateParser(day, row.values[2] || row.values[5]);
                                data.to_date = dateParser(day, row.values[3] || row.values[6]);
                            }

                            axios.post(data.date ? vacation : regular,{ data })
                                .then(response => {
                                    console.log(response);
                                })
                                .catch(error => {
                                    console.log(error);
                                });
                        });
                    });
            })
            .catch(error => {
                console.error(error)
            })
    });
};

bootstrap();

