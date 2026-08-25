import React, { Component } from 'react';

class Home extends Component {
    render() {
        return (
            <div className="container-fluid">
                <h1 className="text-center">Aplication for manipulate a Excel Files</h1>
                <p className="text-center">1.- Charge a excel file.</p>
                <p className="text-center">2.- If you charge another file delete the current data.</p>
                <p className="text-center">3.- For export to excel file just press the button.</p>
                <p className="text-center">4.- For export like a SQL INSERT Script just give the name of the table /n
                    and press the button.
                </p>
                <p className="text-center">5.- For export like an Array of Objects in a TXT. Press the button.</p>
            </div>
        );
    };
}

export default Home;