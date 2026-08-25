import React, { Component } from 'react';
import XLSX from 'xlsx';
import FileSaver from 'file-saver';

class ReadFile extends Component {
    constructor(props) {
        super(props);
        let fileName = "completeData.xls";

        if (fileName === undefined || fileName === null) {
            this.state = { allData: [], fileFromExcel: [], headersFileFromExcel: [], fileName: "", errorTxt: true, errorAdd: true, page: 1 };
        } else {
            this.state = { allData: JSON.parse(sessionStorage.getItem(fileName)), fileFromExcel: JSON.parse(sessionStorage.getItem('data')), headersFileFromExcel: JSON.parse(sessionStorage.getItem('headers')), fileName: "completeData.xls", errorTxt: true, errorAdd: true, page: 1 };
        }

    }


    render() {
        let infiniteScroll = "";
        if (this.state.fileFromExcel && this.state.fileFromExcel.length <= 5) {
            infiniteScroll = ""
        } else {
            infiniteScroll = "scrollcase";
        }

        const getFile = (e) => {
            if (e) {
                var file = e.target.files[0];
                if (!file) return;
                var reader = new FileReader();
                reader.onload = function (e) {
                    var data = new Uint8Array(e.target.result);
                    var workbook = XLSX.read(data, { type: 'array' });
                    var firstSheet = workbook.Sheets[workbook.SheetNames[0]];

                    // header: 1 instructs xlsx to create an 'array of arrays'
                    var result = XLSX.utils.sheet_to_json(firstSheet);

                    console.log(result);
                    if (result.length > 0) {
                        //save the array in a the global variable
                        onClick(result, "completeData.xls");
                    } else {
                        window.alert("The file doesn't has more than 1 line.")
                    }


                };


                reader.readAsArrayBuffer(file);
            }

        };


        const onClick = (result, fileName) => {
            sessionStorage.clear();
            let data = [];
            let headers = result[0] ? Object.getOwnPropertyNames(result[0]) : prompt("Format error the file should have mor that 1 line");
            let index = headers[0].indexOf("__rowNum__");
            if (index > -1) {
                headers.splice(index, 1);
            }

            result.forEach((element, i) => {
                data.push(Object.values(element));

            });

            this.setState({
                allData: result,
                fileFromExcel: data,
                headersFileFromExcel: headers
            });

            sessionStorage.setItem('headers', JSON.stringify(headers));
            sessionStorage.setItem('data', JSON.stringify(data));
            sessionStorage.setItem(fileName, JSON.stringify(result));
            window.location.reload();
        };

        const EditDataFile = () => {
            let fileName = "completeData.xls";
            let data = {};
            let datafileFromExcel = [];
            let dataExistsFile = JSON.parse(sessionStorage.getItem(fileName));
            for (let i = 0; i < this.state.headersFileFromExcel.length; i++) {
                if (typeof this.state.fileFromExcel[0][i] === "number") {
                    data[this.state.headersFileFromExcel[i]] = parseInt(document.getElementById(`${this.state.headersFileFromExcel[i]}`).value);
                } else {
                    data[this.state.headersFileFromExcel[i]] = document.getElementById(`${this.state.headersFileFromExcel[i]}`).value;
                }
            }

            dataExistsFile.push(data);

            dataExistsFile.forEach((element, i) => {
                datafileFromExcel.push(Object.values(element));

            });


            sessionStorage.setItem('data', JSON.stringify(datafileFromExcel));
            sessionStorage.setItem(fileName, JSON.stringify(dataExistsFile));

            this.setState({
                allData: JSON.parse(sessionStorage.getItem(fileName)),
                fileFromExcel: JSON.parse(sessionStorage.getItem('data'))
            });

            for (let i = 0; i < this.state.headersFileFromExcel.length; i++) {
                document.getElementById(`${this.state.headersFileFromExcel[i]}`).value = "";
            }
        };

        const helperDynamicData = (data) => {
            let headers = this.state.headersFileFromExcel;
            let obj = {};
            for (let i = 0; i < data.length; i++) {
                obj[headers[i]] = data[i]
            }
            return obj;
        };

        const updateFileData = (event, keyIndex, index, type) => {
            if (event) {
                let fileName = "completeData.xls";
                let data1 = JSON.parse(sessionStorage.getItem("data"));
                let data2 = [];
                if (type === "number") {
                    data1[keyIndex][index] = parseInt(document.getElementById(`td${keyIndex}${index}`).value);
                } else {
                    data1[keyIndex][index] = document.getElementById(`td${keyIndex}${index}`).value;
                }
                for (let i = 0; i < data1.length; i++) {
                    data2.push(helperDynamicData(data1[i]))
                }
                this.setState({
                    allData: data2,
                    fileFromExcel: data1

                });
                sessionStorage.setItem("data", JSON.stringify(data1));
                sessionStorage.setItem(fileName, JSON.stringify(data2));
            }
        };

        const becomeExcelFile = () => {
            let fileName = "completeData.xls".split('.');
            let fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
            let fileNameNormal = "completeData.xls";
            let allData = JSON.parse(sessionStorage.getItem(fileNameNormal));
            const ws = XLSX.utils.json_to_sheet(allData);
            const wb = { Sheets: { 'data': ws }, SheetNames: ['data'] };
            const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            const data = new Blob([excelBuffer], { type: fileType });
            FileSaver.saveAs(data, fileName[0]);
        };

        const becomeSQLFile = () => {
            let table = document.querySelector('#table').value;
            let fileName = "completeData.xls";
            let fileNameSplit = "completeDataSQL.xls".split('.');
            let allData = JSON.parse(sessionStorage.getItem(fileName));
            let headers = Object.getOwnPropertyNames(allData[0]);
            let fileType = 'text/plain;charset=UTF-8';
            let data1 = '';
            for (var i = 0; i < allData.length; i++) {
                data1 = data1 + `INSERT INTO ${table} (${headers.join()}) VALUES (${JSON.stringify(Object.values(allData[i])).replace("[", "").replace("]", "")});\n`;
            }
            const file = new Blob([data1], { type: fileType });
            FileSaver.saveAs(file, fileNameSplit[0]);
            document.querySelector('#table').value = "";
        };

        const becomeArrayFile = () => {
            let fileName = "completeData.xls";
            let fileNameSplit = "completeDataArray.xls".split('.');
            let allData = sessionStorage.getItem(fileName);
            let fileType = 'text/plain;charset=UTF-8';
            const file = new Blob([allData], { type: fileType });
            FileSaver.saveAs(file, fileNameSplit[0]);
        };

        const errorHandlerTxt = (e) => {
            if (e) {
                let table = e.target.value;
                if (table.length >= 1) {
                    this.setState({
                        errorTxt: false
                    })
                } else {
                    this.setState({
                        errorTxt: true
                    })
                }
            }


        };



        const errorHandlerAdd = (e) => {
            if (e) {
                let data = [];
                for (let i = 0; i < this.state.headersFileFromExcel.length; i++) {
                    if (document.getElementById(`${this.state.headersFileFromExcel[i]}`).value.length >= 1) {
                        data.push(i);
                    }
                }
                if (data.length === this.state.headersFileFromExcel.length) {
                    this.setState({
                        errorAdd: false
                    })
                } else {
                    this.setState({
                        errorAdd: true
                    })
                }
            }
        };

        const dataTable = (page) => {
            let quantityData = 5 * page;
            let keyArray = [];


            for (let i = 0; i < quantityData; i++) {
                keyArray.push(i)
            }

            return keyArray.slice(-5).map((key) => {
                return infiniteData(key, page)
            })
        };

        const infiniteData = (key, page) => {
            if (this.state.fileFromExcel[key] !== undefined) {
                return <tr key={key}>
                    {this.state.fileFromExcel[key].map((dataElement, i) => {
                        if (dataElement && typeof dataElement === "number") {
                            return <td key={i}><input key={i} type="number" className="form-control none-border" id={`td${key}${i}`} aria-describedby={dataElement} name={`td${key}${i}`} placeholder={`Insert the data`} defaultValue={dataElement} onChange={(e) => { updateFileData(e, key, i, "number") }} /></td>
                        } else {
                            return <td key={i}><input key={i} type="text" className="form-control none-border" id={`td${key}${i}`} aria-describedby={dataElement} name={`td${key}${i}`} placeholder={`Insert the data`} defaultValue={dataElement} onChange={(e) => { updateFileData(e, key, i, "string") }} /></td>
                        }

                    })}
                </tr>
            }
        };

        const nextPage = (event, page) => {
            if (event) {
                let nextP = page + 1;
                let maxPage = Math.ceil(this.state.fileFromExcel.length / 5);
                if (nextP <= maxPage) {
                    this.setState({ page: this.state.page + 1 });
                } else {
                    this.setState({ page: 1 });
                }
            }
        };

        const previousPage = (event, page) => {
            if (event) {
                if (page === 1) {
                    this.setState({
                        page: Math.ceil(this.state.fileFromExcel.length / 5)
                    });
                } else {
                    this.setState({
                        page: this.state.page - 1
                    });
                }
            }
        };

        const formToData = (item, i) => {
            return <div className="mb-3" id={i} key={i}>
                <label className="form-label text-capitalize">{item}:</label>
                {typeof this.state.fileFromExcel[0][i] === "number" ?
                    <input type="number" className="form-control formData" id={item} aria-describedby={item} key={i} name={item} placeholder={`Insert the ${item}`} required onChange={(e) => { errorHandlerAdd(e) }} />
                    :
                    <input type="text" className="form-control formData" id={item} aria-describedby={item} key={i} name={item} placeholder={`Insert the ${item}`} required onChange={(e) => { errorHandlerAdd(e) }} />
                }

            </div>


        };



        return (
            <div className="container mt-3">
                {this.state.allData && this.state.allData.length >= 1 ?
                    <div>
                        <h2 className="text-center">Now you can:</h2>
                        <div className="d-flex justify-content-center mt-3">
                            <button type="button" className="btn btn-primary me-2" data-bs-toggle="modal" data-bs-target="#exampleModal">Edit a File</button>
                            <button type="button" className="btn btn-success me-2" onClick={becomeExcelFile}>Become in Excel</button>
                            <button type="button" className="btn btn-primary me-2" data-bs-toggle="modal" data-bs-target="#exampleModal2">SQL Insert Script in a TXT</button>
                            <button type="button" className="btn btn-primary me-2" onClick={becomeArrayFile}>Array in a TXT</button>
                            {/* Modal Edit  */}
                            <div className="modal fade" id="exampleModal" aria-labelledby="exampleModalLabel" aria-hidden="true">
                                <div className="modal-dialog">
                                    <div className="modal-content">
                                        <div className="modal-header">
                                            <h5 className="modal-title" id="exampleModalLabel">Data to insert in the file:</h5>
                                            <button type="button" className="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                                        </div>
                                        <div className="modal-body">
                                            {this.state.headersFileFromExcel.map((item, i) => (
                                                formToData(item, i)
                                            ))}
                                        </div>
                                        <div className="modal-footer">
                                            <button type="button" className="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                                            {this.state.errorAdd === true ?
                                                <button type="button" className="btn btn-primary" disabled>Insert the data</button>
                                                :
                                                <button type="button" className="btn btn-primary" onClick={EditDataFile}>Insert the data</button>
                                            }
                                        </div>
                                    </div>
                                </div>
                            </div>

                            {/* Modal Txt  */}
                            <div className="modal fade" id="exampleModal2" aria-labelledby="exampleModalLabel" aria-hidden="true">
                                <div className="modal-dialog">
                                    <div className="modal-content">
                                        <div className="modal-header">
                                            <h5 className="modal-title" id="exampleModalLabel">Data to insert in the file:</h5>
                                            <button type="button" className="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                                        </div>
                                        <div className="modal-body">
                                            {
                                                <div className="mb-3">
                                                    <label className="form-label text-capitalize">Table to insert:</label>
                                                    <input type="text" className="form-control formData" id="table" aria-describedby="table" name="table" placeholder={`Insert the Table in SQL`} required onChange={(e) => { errorHandlerTxt(e) }} />
                                                </div>
                                            }
                                        </div>
                                        <div className="modal-footer">
                                            <button type="button" className="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                                            {this.state.errorTxt === true ?
                                                <button type="button" className="btn btn-primary" disabled>Generic a txt with the INSERT SQL Script</button>
                                                :
                                                <button type="button" className="btn btn-primary" data-bs-dismiss="modal" onClick={becomeSQLFile}>Generic a txt with the INSERT SQL Script</button>
                                            }
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    :
                    <span></span>
                }
                <h3 className="form-label mt-3 mb-4">Put your excel to manage:</h3>
                <input className="form-control mb-4" type="file" id="formFile" onChange={(e) => { getFile(e) }} />
                {this.state.fileFromExcel && this.state.fileFromExcel.length >= 1 ?
                    <span className="mb-4">Page {this.state.page} of {Math.ceil(this.state.fileFromExcel.length / 5)}</span>
                    : <span></span>}
                {this.state.fileFromExcel && this.state.fileFromExcel.length >= 1 ?
                    <div className={infiniteScroll} >
                        <table className="table table-hover mt-4">
                            <thead>
                                <tr>
                                    {this.state.headersFileFromExcel.map((item, i) => (
                                        <th scope="col" className="text-capitalize" key={i}>{item}</th>
                                    ))}
                                </tr>
                            </thead>
                            <tbody >
                                {dataTable(this.state.page)}

                            </tbody>
                        </table>
                    </div>
                    :
                    <div>
                        <h3 className="mt-3">No excel readed</h3>
                    </div>
                }
                {this.state.fileFromExcel && this.state.fileFromExcel.length >= 1 ?
                    <nav aria-label="Page navigation example">
                        <ul className="pagination justify-content-center mt-4 mb-4">
                            <li className="page-item">
                                <button className="page-link" onClick={(e) => { if (e) this.setState({ page: 1 }) }}>First</button>
                            </li>
                            <li className="page-item">
                                <button className="page-link" onClick={(e) => { previousPage(e, this.state.page) }}>Previous</button>
                            </li>
                            <li className="page-item">
                                <button className="page-link" onClick={(e) => { nextPage(e, this.state.page) }}>Next</button>
                            </li>
                            <li className="page-item">
                                <button className="page-link" onClick={(e) => { if (e) this.setState({ page: Math.ceil(this.state.fileFromExcel.length / 5) }) }}>Last</button>
                            </li>
                        </ul>
                    </nav>
                    :
                    <span></span>
                }
            </div>

        );
    };
}

export default ReadFile;