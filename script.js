document.getElementById('fileInput').addEventListener('change', handleFile);

function handleFile(event) {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const transpose = a => a[0].map((_, c) => a.map(r => r[c]));
        const number1 = Math.trunc(document.getElementById("number1").value);

        // "XXX"という名前のシートを取得
        const sheetNames = ["XXX","YYY","ZZZ"];
        const keycodes = ["A-1","A-2","A-3","B-1","B-2","B-3","B-4","B-5"];
        let KeyCount = Array(keycodes.length).fill(0);
        let KeyCount_Bysheet= new Array(sheetNames.length); //要素数5の配列(array)を作成
        for(let y = 0; y < sheetNames.length; y++) {
            KeyCount_Bysheet[y] = new Array(keycodes.length).fill(0); //配列(array)の各要素に対して、要素数5の配列を作成し、0で初期化
        }

        for (let j = 0; j<sheetNames.length; j++) {
            const sheet = workbook.Sheets[sheetNames[j]];
            const tmp_range = sheet['!ref'];

            // セルの範囲 B2～K6 を取得し、その中から"A-1"と"B-3"をカウントする
            const range = XLSX.utils.decode_range(tmp_range);

            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellAddress = { c: C, r: R };
                    const cellRef = XLSX.utils.encode_cell(cellAddress);
                    const cellValue = sheet[cellRef] ? sheet[cellRef].v : null;

                    for (let i = 0; i<keycodes.length; i++) {
                        if (cellValue === keycodes[i]) {
                            KeyCount[i]++;
                            KeyCount_Bysheet[j][i]++;
                            break;
                        }
                    }
                }
            }

        }
        // 棒グラフを表示
        displayBarChart(KeyCount, keycodes, number1);
        console.log(KeyCount_Bysheet);
        displayStackChart1(KeyCount_Bysheet, keycodes, sheetNames, number1);
        const KeyCount_Bysheet_t = transpose(KeyCount_Bysheet);
        displayStackChart2(KeyCount_Bysheet_t, keycodes, sheetNames, number1);
    };

    reader.readAsArrayBuffer(file);
}

// Chart.jsを使って棒グラフを描画
function displayBarChart(KeyCount, keycodes, number1) {
    const ctx = document.getElementById('myChart1').getContext('2d');
    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: keycodes,
            datasets: [{
                label: '合計',
                data: KeyCount,
                //backgroundColor: ['#FF6384', '#36A2EB'],
                //borderColor: ['#FF6384', '#36A2EB'],
                borderWidth: 2
            }]
        },
        options: {
            plugins: {
                title: {
                    display: true,
                    text: 'コード別合計',
                    font: {
                        size: 18,
                    }
                }
            },
            y: {
                min: 0, //Y軸の最小値
                max: number1, //Y軸の最大値
            },
            scales: {
                x: {
                    ticks: {
                        color: "blue",
                        font: {
                            size: 18
                        }
                    }
                },
                y: {
                    ticks: {
                        font: {
                            size: 18
                        }
                    },
                    title: {
                        display: true,
                        text: '時間（ｈ）',//Y軸のﾗﾍﾞﾙ
                        font: {
                            size: 18
                        }
                    },
                    beginAtZero: true
                }
            }
        }
    });
}

function displayStackChart1(KeyCount_Bysheet, keycodes, sheetNames, number1) {
    const ctx = document.getElementById('myChart2').getContext('2d');
    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: keycodes,
            datasets: [
                {
                    label: sheetNames[0],
                    data: KeyCount_Bysheet[0],
                    //backgroundColor: ['#FF6384', '#36A2EB'],
                    //borderColor: ['#FF6384', '#36A2EB'],
                    borderWidth: 2
                },
                {
                    label: sheetNames[1],
                    data: KeyCount_Bysheet[1],
                    borderWidth: 2
                },
                {
                    label: sheetNames[2],
                    data: KeyCount_Bysheet[2],
                    borderWidth: 2
                }
            ]
        },
        options: {
            plugins: {
                title: {
                    display: true,
                    text: 'コード別内訳',
                    font: {
                        size: 18,
                    }
                }
            },
            y: {
                min: 0, //Y軸の最小値
                max: number1, //Y軸の最大値
            },
            scales: {
                x: {
                    ticks: {
                        color: "blue",
                        font: {
                            size: 18
                        }
                    },
                    stacked: true,
                },
                y: {
                    ticks: {
                        font: {
                            size: 18
                        }
                    },
                    title: {
                        display: true,
                        text: '時間（ｈ）',//Y軸のﾗﾍﾞﾙ
                        font: {
                            size: 18
                        }
                    },
                    stacked: true,
                    beginAtZero: true
                }
            }
        }
    });
}

function displayStackChart2(KeyCount_Bysheet, keycodes, sheetNames, number1) {
    const ctx = document.getElementById('myChart3').getContext('2d');
    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sheetNames,
            datasets: [
                {
                    label: keycodes[0],
                    data: KeyCount_Bysheet[0],
                    borderWidth: 2
                },
                {
                    label: keycodes[1],
                    data: KeyCount_Bysheet[1],
                    borderWidth: 2
                },
                {
                    label: keycodes[2],
                    data: KeyCount_Bysheet[2],
                    borderWidth: 2
                },
                {
                    label: keycodes[3],
                    data: KeyCount_Bysheet[3],
                    borderWidth: 2
                },
                {
                    label: keycodes[4],
                    data: KeyCount_Bysheet[4],
                    borderWidth: 2
                },
                {
                    label: keycodes[5],
                    data: KeyCount_Bysheet[5],
                    borderWidth: 2
                },
                {
                    label: keycodes[6],
                    data: KeyCount_Bysheet[6],
                    borderWidth: 2
                },
                {
                    label: keycodes[7],
                    data: KeyCount_Bysheet[7],
                    borderWidth: 2
                }
            ]
        },
        options: {
            plugins: {
                title: {
                    display: true,
                    text: '人別内訳',
                    font: {
                        size: 18,
                    }
                }
            },
            y: {
                min: 0, //Y軸の最小値
                max: number1, //Y軸の最大値
            },
            scales: {
                x: {
                    ticks: {
                        color: "blue",
                        font: {
                            size: 18
                        }
                    },
                    stacked: true,
                },
                y: {
                    ticks: {
                        font: {
                            size: 18
                        }
                    },
                    title: {
                        display: true,
                        text: '時間（ｈ）',//Y軸のﾗﾍﾞﾙ
                        font: {
                            size: 18
                        }
                    },
                    stacked: true,
                    beginAtZero: true
                }
            }
        }
    });
}
