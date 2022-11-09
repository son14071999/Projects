import { Component, OnInit } from '@angular/core';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit{
  ngOnInit(): void {
    const translate = require("../../node_modules/translate")
    translate.from = "en"
    translate.to = "vi"
    for(let i = 0; i <= 10000; i++) {
      setTimeout(() => {
        translate("Hello").then((resp: any) => {
          console.log(i)
          console.log(resp)
        })
      }, 20);
    }
  }
  title = 'app-voca';
}
