# The only auto clicker you will ever need

Has all the features you could ask for:  

* Static CPS & +/- CPS variance or Manual time definition.
* Can press any mouse or keyboard button
* Hold (press), toggle and toggle (with seperate stop bind) modes
* Manual start and stop binds
* Force stop key

> ## For developers
>
> script can be compiled with the helper script:
>
> ```bash
> python compile_nuitka.py
> ```
>
> it hashes an existing exe, runs Nuitka, captures output to `build/nuitka-build.log`,
> and verifies the output exe changed.
>
> (or use `pyinstaller clicker.py --onefile --noconsole -n=TheBestAutoClickerOAT` if using pyinstaller)
>
> Additionally you can compile an installer using [Inno Setup](https://jrsoftware.org/isinfo.php) and `TheBestAutoClickerOAT.iss`
