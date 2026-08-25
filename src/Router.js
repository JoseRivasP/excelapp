import React, { Component } from 'react';
import { BrowserRouter, Route, Switch } from 'react-router-dom';
import Navbar from './components/Navbar/Navbar';
import Home from './components/Home/Home';
import ReadFile from './components/ReadFile/ReadFile';

class Router extends Component {
    componentDidUpdate() {
        let fileNameStorage = Object.keys(sessionStorage)[0];
        if (JSON.parse(sessionStorage.getItem(fileNameStorage)) >= 1) {
            let data = [];
            let headers = Object.getOwnPropertyNames(JSON.parse(sessionStorage.getItem(this.state.fileName))[0]);
            JSON.parse(sessionStorage.getItem(fileNameStorage)).forEach((element, i) => {
                data.push(Object.values(element));

            });
            this.setState({ allData: JSON.parse(sessionStorage.getItem(fileNameStorage)), fileFromExcel: data, headersFileFromExcel: headers, fileName: fileNameStorage });
        }
    }
    render() {
        return (
            <BrowserRouter>
                <Navbar />
                <Switch>
                    <Route exact path="/" component={Home} />
                    <Route exact path="/readFile" component={ReadFile} />
                </Switch>
            </BrowserRouter>
        );
    };
}

export default Router;