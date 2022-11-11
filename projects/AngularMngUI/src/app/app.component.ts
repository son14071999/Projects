import { Component, OnInit } from '@angular/core';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {
  ngOnInit(): void {
  }
  title = 'AngularMngUI';
  logoLogin = "https://ra-dev.mobica.vn/images/loginLogo.png";
  icon = "https://ra-dev.mobica.vn/images/headerLogo.png"
}
