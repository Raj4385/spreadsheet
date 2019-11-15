import React, { Component } from 'react';
import { Parser, ERROR_REF, error as isFormulaError } from 'hot-formula-parser';
import './App.css';

class App extends Component {

  constructor(props) {
    super(props);
    this.state = {
      data: null,
      dataArray: [],
      selectedCell: '',
      formula: null
    };
  }

  componentWillMount() {
    var l_data = require('../src/data/cmbs_deal_model.json');
    var l_formula = require('../src/data/formulas.json');
    var l_dataArray = [];
    l_dataArray[0] = {}

    var header = 0;
    for (var key of Object.keys(l_data.DealData[0].LoanData[0])) {
      var secondLabel = String.fromCharCode(0 % 26 + 0x41);
      var firstLabel = (0 / 26 > 0 && 0 > 25) ? String.fromCharCode(((0 / 26) - 1) % 26 + 0x41) : '';
      var column = firstLabel + secondLabel + header.toString();
      l_dataArray[0][column] = key;
      header += 1;
    }

    l_data.DealData[0].LoanData.forEach((el, i) => {
      var j = 0;
      var index = i + 1;
      l_dataArray[index] = {}
      for (var key of Object.keys(el)) {
        var secondLabel = String.fromCharCode(index % 26 + 0x41);
        var firstLabel = (index / 26 > 0 && index > 25) ? String.fromCharCode(((index / 26) - 1) % 26 + 0x41) : '';
        var column = firstLabel + secondLabel + j.toString();
        l_dataArray[index][column] = el[key];
        j += 1;
      }
    });
    this.setState({ data: l_dataArray, formula: l_formula });
  }

  getCellId(e) {
    console.log(e.target.id);
    this.setState({ selectedCell: e.target.id });
  }

  createTable = () => {
    let table = []
    var parser = new Parser();

    // Outer loop to create parent
    for (let j = 0; j < Object.keys(this.state.data[0]).length; j++) {
      let children = []
      //Inner loop to create children
      for (let i = 0; i < this.state.data.length; i++) {
        var secondLabel = String.fromCharCode(i % 26 + 0x41);
        var firstLabel = (i / 26 > 0 && i > 25) ? String.fromCharCode(((i / 26) - 1) % 26 + 0x41) : '';
        var column = firstLabel + secondLabel + j;
        var rowFormula = (this.state.formula[column] != undefined ? this.state.formula[column] : null);
        var Result = null;
        if (rowFormula != null) {
          var obj = parser.parse(this.MatchAndReplace(rowFormula.replace(/ /g, '')));
          Result = obj.error == null ? obj.result : obj.error;
        }
        children.push(<td id={column} style={rowFormula != null ? { border: '1px solid black', backgroundColor: 'yellow' } : { border: '1px solid black' }} onClick={(e) => { this.getCellId(e) }}>{Result != null ? Result : this.state.data[i][column]}</td>)
      }
      //Create the parent and add the children
      table.push(<tr style={{ border: '1px solid black' }}>{children}</tr>)
    }
    return table
  }

  MatchAndReplace(formula) {
    if (formula.indexOf(':') != -1) {
      var findcell = this.ParseAllCellAddressForRange(formula)
      for (let m = 0; m < findcell.length - 1; m++) {
        if (formula.indexOf(findcell[m] + ':' + findcell[m + 1])) {
          var lineItem1 = this.SeparteCellNumberAndName(findcell[m])
          var lineItem2 = this.SeparteCellNumberAndName(findcell[m + 1])
          if (lineItem1.Name == lineItem2.Name)
            formula = this.ModifyFormulaColumnRange(lineItem1.Number, lineItem2.Number, lineItem1.Name, formula, findcell[m], findcell[m + 1]);
          else if (lineItem1.Number == lineItem2.Number)
            formula = this.ModifyFormulaRowRange(lineItem1.Name, lineItem2.Name, lineItem1.Number, formula, findcell[m], findcell[m + 1]);
        }
      }
    }

    for (let i = 0; i < this.state.data.length; i++) {
      var Keys = Object.keys(this.state.data[i]);
      for (let j = Keys.length - 1; j > 0; j--) {
        if (formula.indexOf(Keys[j]) !== -1) {
          var regex = new RegExp(Keys[j], "gi")
          formula = formula.replace(regex, this.state.data[i][Keys[j]]);
        }
      }
    }

    return formula;
  }

  ParseAllCellAddressForRange(formula) {
    var specialChars = [',', ')', '(', ':'];
    var specialRegex = new RegExp('[' + specialChars.join('') + ']');
    var strformula = formula.split(specialRegex);
    var findcell = []
    for (let k = 0; k < strformula.length; k++) {
      for (let i = 0; i < this.state.data.length; i++) {
        var Keys = Object.keys(this.state.data[i]);
        for (let j = Keys.length - 1; j > 0; j--) {
          if (strformula[k].trim() == Keys[j]) {
            findcell.push(strformula[k].trim());
          }
        }
      }
    }
    return findcell;
  }

  SeparteCellNumberAndName(findcell) {
    var cell1Name = '';
    var cell1Number = '';

    for (let l = 0; l < findcell.length; l++) {
      if (findcell.charAt(l) >= 0 && findcell.charAt(l) <= 9)
        cell1Number += findcell.charAt(l)
      else
        cell1Name += findcell.charAt(l)
    }
    var line = {
      "Name": cell1Name,
      "Number": cell1Number
    }
    return line;
  }

  ModifyFormulaColumnRange(cell1Number, cell2Number, cell1Name, formula, findcell, findNextcell) {
    var Min = cell1Number < cell2Number ? cell1Number : cell2Number;
    var Max = cell1Number > cell2Number ? cell1Number : cell2Number;
    var expandFormaula = '';
    for (let i = Min; i < Max; i++) {
      expandFormaula += cell1Name + i.toString() + ',';
    }
    expandFormaula += cell1Name + Max.toString();
    formula = formula.replace(findcell + ':' + findNextcell, expandFormaula);
    return formula;
  }

  ModifyFormulaRowRange(cell1Name, cell2Name, cell1Number, formula, findcell, findNextcell) {
    var inNum1 = this.GetCellNameInCharCode(cell1Name, cell1Name.length - 1)
    var inNum2 = this.GetCellNameInCharCode(cell2Name, cell2Name.length - 1)

    var Min = inNum1 < inNum2 ? inNum1 : inNum2;
    var Max = inNum1 > inNum2 ? inNum1 : inNum2;

    formula = this.ResolveCellAddressForRowRange(cell1Number, findcell, findNextcell, formula, Min, Max)
    return formula;
  }

  GetCellNameInCharCode(cell1Name, len) {
    var inNum = '';
    for (let i = 0; i < cell1Name.length; i++) {
      if (i == 0) {
        inNum = cell1Name.charAt(len).charCodeAt(0) - 65;
      }
      else {
        inNum += (Math.pow(26, i) * (((cell1Name.charAt(len).charCodeAt(0)) - 65) + 1));
      }
      len -= 1
    }
    return inNum;
  }

  ResolveCellAddressForRowRange(cell1Number, findcell, findnextcell, formula, Min, Max) {
    var expandFormaula = '';
    for (let i = Min; i <= Max; i++) {
      var value = i;
      var iteration = 1;
      var LimitReached = false;
      while (LimitReached == false) {
        if (value < Math.pow(26, iteration)) {
          LimitReached = true;
          var str = '';
          for (let i = iteration - 1; i >= 0; i--) {
            var Resolved = (i == 0) ? value : (Math.floor(value / Math.pow(26, i)) - 1);
            str += String.fromCharCode(Resolved + 0x41);
            value = value - (Math.pow(26, i) * (Resolved + 1));
          }
          expandFormaula += str + cell1Number + ",";
        }
        else
          iteration++;
      }
    }
    formula = formula.replace(findcell + ':' + findnextcell, expandFormaula.substring(0, expandFormaula.length - 1));
    return formula;
  }

  render() {
    return (
      <div >
        <table style={{ border: '1px solid black' }}>
          {this.createTable()}
        </table>
      </div>
    );
  }
}

export default App;
