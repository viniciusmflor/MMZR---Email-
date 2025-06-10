import { Component } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { EmailGeneratorComponent } from './components/email-generator/email-generator.component';

@Component({
  selector: 'app-root',
  imports: [RouterOutlet, EmailGeneratorComponent],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  title = 'mmzr-email-generator';
}
