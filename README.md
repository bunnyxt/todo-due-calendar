# ToDo Due Calendar

Turning unfinished Microsoft ToDo tasks with due date into ICS calendar feed.

# Getting Started

1. Go to Azure, search for Microsoft Entra ID service, register a new app (or use existing one). Write down client ID (Application ID).
2. Clone this repo, install dependencies using `pnpm install`.
3. Create `.env` from `.env.example` file, fill in `CLIENT_ID` we got from step 1. Also, set `ICS_TOKEN` to a random string. Keep it secret.
4. Run `pnpm run get-tokens` and follow the instructions: open link, paste code, login with your account, allow access. Once finished successfully, you will see `ACCESS_TOKEN` and `REFRESH_TOKEN` printed in the console. Update both in `.env` file.
5. Run `pnpm run dev` to start web app. Your feed will start on `http://localhost:3000/todo-due.ics?token=your-token`.
6. For debugging and testing, manually import the generated ics file into outlook desktop application.

# Deployment

Fork this repo in GitHub, connect to Vercel, add environment variables there, then you will get a public URL like `https://your-todo-due-calendar.vercel.app/todo-ics?token=your-token`. Go to [outlook](https://outlook.live.com/calendar) website, click "Add calendar", select "Subscribe from web", then fill in the link above.

**IMPORTANT**: Please use a complex ics token and keep it secret. Anyone with the token can access to your ICS calendar feed that might contains sensitive private information.

Another limitation is that, we **CANNOT** control when outlook will fetch the latest calendar from out feed. Delete and re-add the calendar could be a way but definite;y takes time and inconvenient.

# Contact

[bunnyxt](https://github.com/bunnyxt)

# License

[MIT](LICENSE)
