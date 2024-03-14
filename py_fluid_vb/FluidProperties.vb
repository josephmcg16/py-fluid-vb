Public Class FluidProperties
    Public Function NitrogenPropertiesLimitedRange(T As Double, P As Double) As Double

        '   Function to calculate the density of gaseous nitrogen.

        '   For each isotherm, the reference compression factor Z, calculated using the NIST REFPROP package
        '   (NIST Standard Reference Database 23, Version 9.0) implementation of the reference-quality formulation
        '   of Span, R., Lemmon, E.W., Jacobsen, R.T., Wagner, W. and Yokozeki, A. "A Reference Equation of State
        '   for the Thermodynamic Properties for Nitrogen for Temperatures from 63.151 to 1000 K at Pressures up to 2200 MPa",
        '   J. Phys. Chem. Ref. Data, 29(6):1361-1433, 2000, was fitted to a virial equation in normalised density.
        '   The virial coefficients were fitted to polynomials in normalised temperature.

        '   This implementation uses coefficients with 6 significant digits.

        '   The correlation is valid from 0.95 to 20 bar(a) between 245 and 330 K.  Across its full rnage of validity,
        '   the correlation fits the reference data to within 0.0005 %.  The uncertainty of the reference data is 0.02 % (at k = 2),
        '   so the uncertainty of densities calculated from this correlation will also be 0.02 % (at k = 2).

        '   Inputs
        '       T   -   absolute temperature, in K
        '       P   -   absolute pressure, in Pa
        '       RhoOrZ - set to 0 (by default) to return density, set to 1 to return compressibility.

        '   Output
        '       RhoN2   -   nitrogen gas density, in kg/m3, or compressibility, dimensionless.

        '   Correlation developed by Norman F Glen, November 2020.

        '   Code
        '       Version 1.00
        '       Date    25 November 2020
        '       By      Norman Glen

        '       Version 1.10
        '       Date    26 November 2020
        '       By      Norman Glen
        '       Notes   Input parameter RhoOrZ added, to allow output of density or compressibility.

        '   Correlation limits.
        Const Tmin As Double = 245          ' -28.15°C
        Const Tmax As Double = 330          ' 56.85°C
        Const Pmin As Double = 95000.0#      ' 0.95 bar(a)
        Const Pmax = 2000000.0#               ' 20 bar(a)

        '   Correlation coefficients.
        Dim b(2) As Double
        Dim c(2) As Double

        b(0) = 0.0456609
        b(1) = -0.0417071
        b(2) = -0.0120977

        c(0) = 0.00160988
        c(1) = -0.000770858
        c(2) = 0.000824459

        '   Offset.
        Const Offset As Double = 0.0006    ' Percentage.  To account for changes in the values of universal gas constant and molar mass from NIST values.

        '   Reducing parameters.
        Const Tred As Double = 285      ' In K, approximate mid point of temperature range of correlation.
        Const Rhored As Double = 30     ' In kg/m3, just above maximum density for temperature and pressure range of correlation.

        '   Universal gas constant.
        Const R As Double = 8.31446261815824    ' In J/(kg.mol) - exact numerical value following re-definition of SI base units, 20 May 2019.

        '   Molar mass of nitrogen.
        Const MN2 As Double = 28.01348  ' In g/mol - from NIST database (January 2015).

        '   Specific gas constant.
        Dim Rspec As Double
        Rspec = 1000 * R / MN2          ' In J/(kg.K).

        '   Convergance tolerance.
        Const Tol As Double = 0.0001    ' Percentage.

        Dim Bz As Double, Cz As Double
        Dim Tr As Double, Rhor As Double
        Dim Z As Double
        Dim RhoInit As Double
        Dim RhoInc As Double
        Dim Delta As Double
        Dim RhoN2 As Double

        Dim i As Integer
        Dim SgnDelta1 As Integer, SgnDelta2 As Integer

        If (T >= Tmin) And (T <= Tmax) And (P >= Pmin) And (P <= Pmax) Then

            Tr = T / Tred

            For i = 0 To 2
                Bz = Bz + b(i) / Tr ^ i
            Next i

            For i = 0 To 2
                Cz = Cz + c(i) / Tr ^ i
            Next i

            '       Use 1.05 x ideal gas density as intial estimate
            '       as this will always be higher than the true value
            '       and use 1 % of this value as the initial increment.

            RhoN2 = 1.05 * P / (Rspec * T)

            RhoInc = 0.01 * RhoN2

            RhoInit = RhoN2 - RhoInc

            SgnDelta1 = -1

            Do

                Rhor = RhoInit / Rhored

                Z = 1 + Bz * Rhor + Cz * Rhor ^ 2

                RhoN2 = P / (Rspec * T * Z)

                Delta = 100 * (RhoN2 - RhoInit) / RhoInit

                SgnDelta2 = Math.Sign(Delta)

                If SgnDelta1 <> SgnDelta2 Then
                    SgnDelta1 = SgnDelta2
                    RhoInc = -0.5 * RhoInc
                End If

                RhoInit = RhoInit - RhoInc

            Loop Until Math.Abs(Delta) / Tol < 1

            RhoN2 = RhoN2 / (1 + 0.01 * Offset)
            Z = P / (Rspec * T * RhoN2)
        
            Return RhoN2

        Else
            Return Double.NaN
        End If

    End Function
    Public Function NitrogenIsentropicExponent2019(T As Double, P As Double) As Double
        '   Function to calculate the isentropic exponent of gaseous nitrogen.

        '   For each isotherm, the reference isentropic exponent n, calculated using the NIST REFPROP package
        '   (NIST Standard Reference Database 23, Version 9.0) implementation of the reference-quality formulation
        '   of Span, R., Lemmon, E.W., Jacobsen, R.T, Wagner, W., and Yokozeki, A.,
        '   "A Reference Equation of State for the Thermodynamic Properties of Nitrogen for Temperatures from 63.151 to 1000 K
        '   and Pressures to 2200 MPa", J. Phys. Chem. Ref. Data, 29(6):1361-1433, 2000, was fitted to a virial equation
        '   in normalised density.  The virial coefficients were fitted to polynomials in inverse normalised temperature.

        '   This implementation uses coefficients with 6 significant digits.

        '   The correlation is valid from 1 to 201 bar(a) between 270 and 350 K.  Across its full range of validity,
        '   the correlation fits the reference data to within 0.004 %.  The uncertainty of the reference data is 0.03 % (at k = 2)
        '   at ambient pressure, rising to 0.05 % (at k=2) at the highest pressure so the uncertainty of isnetropic exponents
        '   calculated from this correlation will also be between 0.02 and 0.05 % (at k = 2).

        '   Inputs
        '       T   -   absolute temperature, in K
        '       P   -   absolute pressure, in Pa

        '   Output
        '       IsenN2   -   nitrogen gas isentropic exponent, dimensionless

        '   Correlation developed by Norman F Glen, April 2019.

        '   Code
        '       Version 1.00
        '       Date    17 April 2019
        '       By      Norman Glen

        Try
            '   Correlation limits.
            Const Tmin As Double = 270          ' -3.15°C
            Const Tmax As Double = 350          ' 76.85°C
            Const Pmin As Double = 100000.0#      ' 1 bar(a)
            Const Pmax = 20100000.0#              ' 201 bar(a)

            '   Correlation coefficients.
            Dim a(2) As Double
            Dim b(2) As Double
            Dim c(2) As Double
            Dim d(2) As Double

            a(0) = 1.39383
            a(1) = 0.0000591672
            a(2) = -0.000000117883

            b(0) = 0.461771
            b(1) = -0.074924
            b(2) = -0.0543008

            c(0) = -0.00959131
            c(1) = 0.40506
            c(2) = -0.102794

            d(0) = 0.0276474
            d(1) = -0.0791277
            d(2) = 0.108407

            '   Reducing parameters.
            Const Tred As Double = 310          ' In K, mid point of temperature range of correlation.
            Const Rhored As Double = 250        ' In kg/m3, just above maximum density for temperature and pressure range of correlation.

            Dim Az As Double, Bz As Double, Cz As Double, Dz As Double
            Dim Tr As Double, Rhor As Double ', Rho As Double

            Dim i As Integer

            If (T >= Tmin) And (T <= Tmax) And (P >= Pmin) And (P <= Pmax) Then

                Tr = T / Tred
                Rhor = (NitrogenPropertiesLimitedRange(T, P) - NitrogenPropertiesLimitedRange(T, 100000.0)) / Rhored

                For i = 0 To 2
                    Az = Az + a(i) * T ^ i
                Next i

                For i = 0 To 2
                    Bz = Bz + b(i) / Tr ^ i
                Next i

                For i = 0 To 2
                    Cz = Cz + c(i) / Tr ^ i
                Next i

                For i = 0 To 2
                    Dz = Dz + d(i) / Tr ^ i
                Next i


                NitrogenIsentropicExponent2019 = Az + Bz * Rhor + Cz * Rhor ^ 2 + Dz * Rhor ^ 3
            Else
                NitrogenIsentropicExponent2019 = Double.NaN
            End If
        Catch ex As Exception
            ' Log.Error(ex)
            Return Double.NaN
        End Try


    End Function
    Public Function CStarNitrogenLimitedRange(UpstreamTemperatureABS As Double, UpstreamPressureABS As Double) As Double


        '   Function to calculate C*, critical flow factor, for gaseous nitrogen.

        '   The NIST REFPROP package (NIST Standard Reference Database 23, Version 9.0) was used
        '   to generate the reference values for the development of the correlation.  Values at
        '   atmospheric pressure at 18 temperatures were fitted to a simple linear against a reduced temperature.
        '   The pressure dependence was derived from C* ratios for 503 values calculated from the REFPROP package.

        '   This implementation uses coefficients with 6 significant digits.

        '   The correlation is valid from 0.95 to 20 bar(a) between 245 and 330 K.

        '   Although the REFPROP package does not give a references for its C* calculation
        '   the values generated by it agree to 6 signifcant digits with values given in
        '   ISO 9300:2005, which have a stated uncertainty of 0.1 % (k=2).

        '   Across its full range of validity the correlation fits the NIST reference data
        '   to within 0.010 % except at 92 of the 1044 points, 25 of which have errors of up
        '   to -0.020 %.  Examination of plots of C* against pressure in the range from
        '   95,000 Pa to 120,000 Pa show non-smooth variations of C*, probably due to
        '   numerical differentiation artefacts in the reference calculations for small
        '   pressure steps at low pressures.

        '   Based on the quality of the fits and the claimed uncertainy of the ISO 9300:2005 data,
        '   the uncertainty of C* values calculated from this correlation is assessed to be 0.1 % (at k = 2).

        '   Inputs
        '       T   -   absolute temperature, in K
        '       P   -   absolute pressure, in Pa

        '   Output
        '       CStarN2   -   critical flow factor

        '   Correlation developed by Norman Glen, November 2020.

        '   Code
        '       Version 1.00
        '       Date    03 November 2020
        '       By      Norman Glen

        '       Version 1.10
        '       Date    23 March 2021
        '       By      Norman Glen
        '       Notes   Pressure range extended down from 1 bar(a) to 0.95 bar(a).

        Try
            '   Correlation limits.
            Const Tmin As Double = 245          ' -28.15°C
            Const Tmax As Double = 330          ' 56.85°C
            Const Pmin As Double = 95000.0#       ' 0.95 bar(a)
            Const Pmax = 2000000.0#               ' 20 bar(a)

            '   Correlation coefficients.
            Dim a(1) As Double
            Dim b(4) As Double
            Dim n(2) As Double

            a(0) = 0.68378
            a(1) = 0.00122

            b(1) = -0.000397822
            b(2) = 0.000808753
            b(3) = -0.00000729242
            b(4) = 0.00000806725

            n(1) = 2.0#
            n(2) = -0.8

            '   Reducing parameters.
            Const Tred As Double = 285      ' In K, mid point of temperature range of correlation.

            Dim Tr As Double
            Dim RhoRat As Double
            Dim CStarAmb As Double, CStarRat As Double

            Dim i As Integer

            If (UpstreamTemperatureABS >= Tmin) And (UpstreamTemperatureABS <= Tmax) And (UpstreamPressureABS >= Pmin) And (UpstreamPressureABS <= Pmax) Then

                Tr = Tred / UpstreamTemperatureABS

                RhoRat = NitrogenPropertiesLimitedRange(UpstreamTemperatureABS, UpstreamPressureABS) / NitrogenPropertiesLimitedRange(UpstreamTemperatureABS, 100000.0#) - 1.0#

                For i = 0 To 1
                    CStarAmb = CStarAmb + a(i) * Tr ^ i
                Next i

                CStarRat = 1.0# + (b(1) + b(2) * Tr ^ n(1)) * RhoRat + (b(3) + b(4) * Tr ^ n(2)) * RhoRat ^ 2

                CStarNitrogenLimitedRange = CStarAmb * CStarRat

            Else
                Return Double.NaN
            End If
        Catch ex As Exception
            ' Log.Error(ex)
            Return Double.NaN
        End Try


    End Function
    Public Function CarbonDioxideProperties(TemperatureABS As Double, PressureABS As Double) As Double

        '   Function to calculate the density of gaseous carbon dioxide.

        '   For each isotherm, the reference compression factor Z, calculated using the NIST REFPROP package
        '   (NIST Standard Reference Database 23, Version 9.0) implementation of the reference-quality formulation
        '   of Span, R. and Wagner, W.,"A New Equation of State for Carbon Dioxide Covering the Fluid Region from
        '   the Triple-Point Temperature to 1100 K at Pressures up to 800 MPa"J. Phys. Chem. Ref. Data, 25(6):1509-1596, 1996.,
        '   was fitted to a virial equation in normalised density.
        '   The virial coefficients were fitted to polynomials in inverse normalised temperature.

        '   This implementation uses coefficients with 7 significant digits.

        '   The correlation is valid from 1 to 40 bar(a) between 285 K <= T <= 320 K.  Across its full rnage of validity,
        '   the correlation fits the reference data to within 0.007 %.  The uncertainty of the reference data up to 30 Mpa is up to 0.05 % (at k = 2),
        '   so the uncertainty of densities calculated from this correlation will be 0.0505 % (at k = 2).
        '
        '   The correlation can also be applied between 270 K <= T < 285 K, from 1 bar(a) to 32 bar(a). However, across this range the correlation fits the reference data to within 0.02 %.
        '   The uncertainty of the reference data up to 30 MPa is up to 0.05 % (at k = 2),
        '   so the uncertainty of densities calculated from this correlation between 270 K and 285 K, from 1 bar(a) to 32 bar(a) will be 0.054 % (at k = 2.

        '   Inputs
        '       T   -   absolute temperature, in K
        '       P   -   absolute pressure, in Pa

        '   Output
        '       Density   -   nitrogen gas density, in kg/m3
        '       Z         -   Compressibility Factor

        '   Correlation developed by Gabriele Chinello, January 2023.

        '   Code
        '       Version 1.00
        '       Date    31 January 2023
        '       By      Gabriele Chinello (code adapted from Norman Glen's template)

        Try
            '   Correlation limits.
            Const Tmin As Double = 270          ' -3.15°C
            Const Tmax As Double = 285          ' 12°C
            Const Tmax2 As Double = 320         ' 46.85°C

            Const Pmin As Double = 100000.0#      ' 1 bar(a)
            Const Pmax As Double = 3200000.0#     ' 32 bar(a)
            Const Pmax2 As Double = 4000000.0#    ' 40 bar(a)


            '   Correlation coefficients.
            Dim b(2) As Double
            Dim c(3) As Double
            Dim e(3) As Double


            b(0) = 0.0326088
            b(1) = 0.0273844
            b(2) = -0.3538431


            c(0) = 1.1144587
            c(1) = -3.4226867
            c(2) = 3.5448901
            c(3) = -1.2073063



            e(0) = 0.0000307885
            e(1) = 0
            e(2) = 0
            e(3) = 0

            '   Reducing parameters.
            Const Tred As Double = 304.13      ' In K, mid point of temperature range of correlation.
            Const Rhored As Double = 110    ' In kg/m3, just above maximum density for temperature and pressure range of correlation.

            '   Specific gas constant.
            Const R As Double = 188.9239    ' In J/(kg.K).

            '   Convergance tolerance.
            Const Tol As Double = 0.001     ' Percentage.

            Dim Bz As Double, Cz As Double, Ez As Double
            Dim Tr As Double, Rhor As Double
            Dim Z As Double
            Dim Density As Double
            Dim RhoInit As Double
            Dim RhoInc As Double
            Dim Delta As Double

            Dim i As Integer
            Dim SgnDelta1 As Integer, SgnDelta2 As Integer


            If (TemperatureABS >= Tmin) And (TemperatureABS <= Tmax) And (PressureABS >= Pmin) And (PressureABS <= Pmax) Or (TemperatureABS >= Tmax) And (TemperatureABS <= Tmax2) And (PressureABS >= Pmin) And (PressureABS <= Pmax2) Then

                Tr = TemperatureABS / Tred

                For i = 0 To 2
                    Bz = Bz + b(i) / Tr ^ i
                Next i

                For i = 0 To 3
                    Cz = Cz + c(i) / Tr ^ i
                Next i

                For i = 0 To 3
                    Ez = Ez + e(i) / Tr ^ i
                Next i

                '       Use 1.05 x ideal gas density as intial estimate
                '       as this will always be higher than the true value
                '       and use 1 % of this value as the initial increment.

                Density = 1.05 * PressureABS / (R * TemperatureABS)

                RhoInc = 0.01 * Density

                RhoInit = Density - RhoInc

                SgnDelta1 = -1

                Do

                    Rhor = RhoInit / Rhored

                    Z = 1 + Bz * Rhor + Cz * Rhor ^ 2 + Ez * Rhor ^ 4

                    Density = PressureABS / (R * TemperatureABS * Z)

                    Delta = 100 * (Density - RhoInit) / RhoInit

                    SgnDelta2 = Math.Sign(Delta)

                    If SgnDelta1 <> SgnDelta2 Then
                        SgnDelta1 = SgnDelta2
                        RhoInc = -0.5 * RhoInc
                    End If

                    RhoInit = RhoInit - RhoInc

                Loop Until System.Math.Abs(Delta) / Tol < 1

                Return Density

            Else

                Return Double.NaN

            End If
        Catch ex As Exception
            ' Log.Error(ex)
            Return Double.NaN
        End Try
    End Function
    Public Function CarbonDioxideViscosity(ByVal TemperatureABS As Double, ByVal Density As Double) As Double
            'ARGUMENT No:      1, 2  , 3   , 4
            'DESC: Calculate the viscosity of carbon dioxide
            '-------------------
            'INPUT:
            '(1)       TemperatureABS = Temperature (ITS-90)  , K
            '(2)              Density = Density               , kg/m³
            'OUTPUT:
            '(3)    CarbonDioxideViscosity = Dynamic Viscosity     , Pa.s , Uncert = 1.12% (below 20 bar) and 3.04% (between 20 bar and 40 bar)
            '-------------------
            'NOTES:
            '  1. The implmentation is based on the paper of Laesecke, A. and Muzny, C.D., "Reference Correlation for the Viscosity of Carbon Dioxide,"
            '  J.Phys.Chem.Ref.Data , 46, 13107, 2017, doi: 10.1063/1.4977429.
            '-------------------
            'MODIFICATION HISTORY:
            ' Mod.  Date      By         Description
            '  0    22/04/99  J.Watson   Initial version
            '  1    29/10/00  H G Ryan   Converted to LabView 5.1
            '  2    01/09/04  B.Robson   Converted to VBA
            '  3    22/09/23  G.Chinello Coefficient tuned for gaseous CO2
            '______________________________________
            Try
                Dim c As Double, Boltz As Double, Eps As Double, Sig As Double
                Dim Lns As Double, Omega As Double, Visc0 As Double, Rr As Double
                Dim A1 As Double, A2 As Double, A3 As Double, A4 As Double, A5 As Double
                Dim B1 As Double, B2 As Double, B3 As Double, B4 As Double, B5 As Double
                Dim Term1 As Double, Vresid As Double
                'Initialise
                Dim Visc As Double
                '-------------------
                'Constants
                A1 = 0.62504484238544
                A2 = -0.0238383664665217
                A3 = -0.416441444241634
                A4 = -0.0421532693534214
                A5 = 0.110266160531217
                '
                B1 = -7.11360341003178
                B2 = 4.50897329497442
                B3 = -0.320521038125112
                B4 = -0.00814806224982459
                B5 = -0.0368249982136418
                '-------------------
                'Calculation
                c = 20.442
                Boltz = 1.38062E-23
                Eps = 1.3808483E-21
                Sig = 3.6502496
                'Calculate visc of dilute gas
                Lns = System.Math.Log(TemperatureABS * Boltz / Eps)
                Omega = System.Math.Exp(A1 + Lns * (A2 + Lns * (A3 + Lns * (A4 + Lns * A5))))
                Term1 = (c * TemperatureABS) ^ 0.5
                Visc0 = 0.000003125 * Term1 / (Sig * Sig * Omega)
                'Calculate reduced residual visc
                Rr = Density / 110.0#
                Vresid = B1 / (Rr - B2) + B1 / B2 + Rr * (B3 + Rr * (B4 + B5 * Rr))
                'Calculate total visc
                Visc = Visc0 + 0.000014 * Vresid
                '_________________________
                Return Visc
            Catch ex As Exception
                Return Double.NaN
            End Try



    End Function
    Public Function CarbonDioxideIsentropicExponent(TemperatureABS As Double, PressureABS As Double) As Double

        '   Function to calculate the isentropic exponent of gaseous nitrogen.

        '   For each isotherm, the reference isentropic exponent n,calculated using the NIST REFPROP package (NIST Standard Reference Database 23, Version 9.0)
        '   implementation of the reference-quality formulation of Span, R. and Wagner, W., "A New Equation of State for Carbon Dioxide Covering the Fluid Region
        '   from the Triple-Point Temperature to 1100 K at Pressures up to 800 MPa"J. Phys. Chem. Ref. Data, 25(6):1509-1596, 1996, was fitted to a virial equation
        '   in normalised density.  The virial coefficients were fitted to polynomials in inverse normalised temperature.

        '   This implementation uses coefficients with 6 significant digits.

        '   The correlation is valid from 1 to 40 bar(a) between 285 K <= T <= 320 K.  Across its full rnage of validity,
        '   the correlation fits the reference data to within 0.007 %.  The uncertainty of the reference data is up to XXX % (at k = 2),
        '   so the uncertainty of densities calculated from this correlation will be XXX % (at k = 2).
        '
        '   The correlation can also be applied between 270 K <= T < 285 K, from 1 bar(a) to 32 bar(a. However, across this range the correlation fits the reference data to within 0.02 %.
        '   The uncertainty of the reference data is up to XXX % (at k = 2),
        '   so the uncertainty of densities calculated from this correlation between 270 K and 285 K, from 1 bar(a) to 32 bar(a) will be XXX % (at k = 2.


        '   Inputs
        '       T   -   absolute temperature, in K
        '       P   -   absolute pressure, in Pa

        '   Output
        '       IsenN2   -   nitrogen gas isentropic exponent, dimensionless

        '   Correlation developed by Gabriele Chinello, August 2023.

        '   Code
        '       Version 1.00
        '       Date    30 August 2023
        '       By      Gabriele Chinello (code adapted from Norman Glen's template)

        '   Correlation limits.

        Try
            Const Tmin As Double = 270          ' -3.15°C
            Const Tmax As Double = 285          ' 12°C
            Const Tmax2 As Double = 320         ' 46.85°C

            Const Pmin As Double = 100000.0#      ' 1 bar(a)
            Const Pmax As Double = 3200000.0#     ' 32 bar(a)
            Const Pmax2 As Double = 4000000.0#    ' 40 bar(a)

            '   Correlation coefficients.
            Dim a(2) As Double
            Dim b(2) As Double
            Dim c(5) As Double
            Dim d(5) As Double

            a(0) = 1.520244
            a(1) = -0.001113984
            a(2) = 0.000001118645

            b(0) = 0.10004194
            b(1) = -0.12134598
            b(2) = -0.02888369

            c(0) = 1077.03301055
            c(1) = -5275.23472451
            c(2) = 10328.72209164
            c(3) = -10104.789647
            c(4) = 4939.40077242
            c(5) = -965.09437574


            d(0) = -1139.87552593
            d(1) = 5586.85918063
            d(2) = -10945.02362421
            d(3) = 10712.62227391
            d(4) = -5238.19923491
            d(5) = 1023.6205368



            '   Reducing parameters.
            Const Tred As Double = 304.13          ' In K, mid point of temperature range of correlation.
            Const Rhored As Double = 110           ' In kg/m3, just above maximum density for temperature and pressure range of correlation.

            Dim Az As Double, Bz As Double, Cz As Double, Dz As Double
            Dim Tr As Double, Rhor As Double

            Dim i As Integer

            If (TemperatureABS >= Tmin) And (TemperatureABS <= Tmax) And (PressureABS >= Pmin) And (PressureABS <= Pmax) Or (TemperatureABS >= Tmax) And (TemperatureABS <= Tmax2) And (PressureABS >= Pmin) And (PressureABS <= Pmax2) Then

                Tr = TemperatureABS / Tred
                Rhor = (CarbonDioxideProperties(TemperatureABS, PressureABS) - CarbonDioxideProperties(TemperatureABS, 100000.0#)) / Rhored

                For i = 0 To 2
                    Az = Az + a(i) * TemperatureABS ^ i
                Next i

                For i = 0 To 2
                    Bz = Bz + b(i) / Tr ^ i
                Next i

                For i = 0 To 5
                    Cz = Cz + c(i) / Tr ^ i
                Next i

                For i = 0 To 5
                    Dz = Dz + d(i) / Tr ^ i
                Next i


                CarbonDioxideIsentropicExponent = Az + Bz * Rhor + Cz * Rhor ^ 2 + Dz * Rhor ^ 3

            Else
                Return Double.NaN
            End If
        Catch ex As Exception
            ' Log.Error(ex)
            Return Double.NaN
        End Try
    End Function
    Public Function MethaneProperties(T As Double, P As Double) As Double

        '   Function to calculate the density and compressibility of gaseous methane.

        '   For each isotherm, the reference compression factor Z, calculated using the NIST REFPROP package
        '   (NIST Standard Reference Database 23, Version 9.0) implementation of the reference-quality formulation
        '   of Setzmann, U and Wagner, W. "A New Equation of State and Tables of Thermodynamic Properties for Methane
        '   Covering the range from the Melting Line to 625 K at Pressures up to 1000 MPa",
        '   J. Phys. Chem. Ref. Data, 20(6):1061-1151, 1991, was fitted to a virial equation in normalised density.
        '   The virial coefficients were fitted to polynomials in normalised temperature.

        '   This implementation uses coefficients with 6 significant digits.

        '   The correlation is valid from 0.95 to 20 bar(a) between 245 and 330 K.  Across its full rnage of validity,
        '   the correlation fits the reference data to within 0.002 %.  The uncertainty of the reference data is 0.03 % (at k = 2),
        '   so the uncertainty of densities calculated from this correlation will also be 0.03 % (at k = 2).

        '   Inputs
        '       T   -   absolute temperature, in K
        '       P   -   absolute pressure, in Pa

        '   Output
        '       RhoCH4   -   methane gas density, in kg/m3

        '   Correlation developed by Norman F Glen, July 2020.

        '   Code
        '       Version 1.00
        '       Date    29 July 2020
        '       By      Norman Glen


        Try
            '   Correlation limits.
            Const Tmin As Double = 245          ' -28.15°C
            Const Tmax As Double = 330          ' 56.85°C
            Const Pmin As Double = 95000.0#      ' 0.95 bar(a)
            Const Pmax = 2000000.0#               ' 20 bar(a)

            '   Correlation coefficients.
            Dim b(2) As Double
            Dim c(2) As Double

            b(0) = 0.058096
            b(1) = -0.0853743
            b(2) = -0.0327984

            c(0) = 0.000790655
            c(1) = 0.00174969
            c(2) = 0.00150852

            '   Offset.
            Const Offset As Double = -0.0013    ' Percentage.  To account for changes in the values of universal gas constant and molar mass from NIST values.

            '   Reducing parameters.
            Const Tred As Double = 285      ' In K, approximate mid point of temperature range of correlation.
            Const Rhored As Double = 20     ' In kg/m3, just above maximum density for temperature and pressure range of correlation.

            '   Universal gas constant.
            Const R As Double = 8.31446261815824    ' In J/(kg.mol) - exact numerical value following re-definition of SI base units, 20 May 2019.

            '   Molar mass of methane.
            Const MCH4 As Double = 16.042499    ' In g/mol - from NIST database (January 2015) for molar masses and isotopic compositions for carbon and hydrogen.

            '   Specific gas constant.
            Dim Rspec As Double
            Rspec = 1000 * R / MCH4         ' In J/(kg.K).

            '   Convergance tolerance.
            Const Tol As Double = 0.0001     ' Percentage.

            Dim Bz As Double, Cz As Double
            Dim Tr As Double, Rhor As Double
            Dim Z As Double
            Dim RhoCH4 As Double
            Dim RhoInit As Double
            Dim RhoInc As Double
            Dim Delta As Double

            Dim i As Integer
            Dim SgnDelta1 As Integer, SgnDelta2 As Integer

            If (T >= Tmin) And (T <= Tmax) And (P >= Pmin) And (P <= Pmax) Then

                Tr = T / Tred

                For i = 0 To 2
                    Bz = Bz + b(i) / Tr ^ i
                Next i

                For i = 0 To 2
                    Cz = Cz + c(i) / Tr ^ i
                Next i

                '       Use 1.05 x ideal gas density as intial estimate
                '       as this will always be higher than the true value
                '       and use 1 % of this value as the initial increment.

                RhoCH4 = 1.05 * P / (Rspec * T)

                RhoInc = 0.01 * RhoCH4

                RhoInit = RhoCH4 - RhoInc

                SgnDelta1 = -1

                Do

                    Rhor = RhoInit / Rhored

                    Z = 1 + Bz * Rhor + Cz * Rhor ^ 2

                    RhoCH4 = P / (Rspec * T * Z)

                    Delta = 100 * (RhoCH4 - RhoInit) / RhoInit

                    SgnDelta2 = System.Math.Sign(Delta)

                    If SgnDelta1 <> SgnDelta2 Then
                        SgnDelta1 = SgnDelta2
                        RhoInc = -0.5 * RhoInc
                    End If

                    RhoInit = RhoInit - RhoInc

                Loop Until System.Math.Abs(Delta) / Tol < 1

                RhoCH4 = RhoCH4 / (1 + 0.01 * Offset)
                Z = P / (Rspec * T * RhoCH4)

                Return RhoCH4
            Else

                Return Double.NaN

            End If
        Catch ex As Exception
            ' Log.Error(ex)
            Return Double.NaN
        End Try
    End Function
    Public Function MethaneViscosity(T As Double, P As Double) As Double

        '   Function to calculate the viscosity of gaseous methane.

        '   A simplified form of the NEL equation developed by Watson for methane was used to fit viscosity at ambient pressure
        '   ('Viscosity of Gaaes in Metric Units', National Engineering Laboratory, HMSO, 1972).  The pressure dependance was
        '   derived from viscosity ratios for 864 values calculated using the NIST REFPROP package
        '   (NIST Standard Reference Database 23, Version 9.0) implementation of the reference-quality formulation of
        '   Setzmann, U and Wagner, W. "A New Equation of State and Tables of Thermodynamic Properties for Methane
        '   Covering the range from the Melting Line to 625 K at Pressures up to 1000 MPa",
        '   J. Phys. Chem. Ref. Data, 20(6):1061-1151, 1991009.

        '   This implementation uses coefficients with 6 significant digits.

        '   The correlation is valid from 0.95 to 20 bar(a) between 245 and 330 K.

        '   The uncertainty in the NIST reference data is 0.3 % (at k=2) and the correlation implemented
        '   in this function fits these data to within 0.003 %, so the certainty of viscosities calculated from the function
        '   at ambient pressure will also be 0.3 % (at k=2).

        '   The NIST reference data were used to determmine the viscosity ratio at each isotherm and a function derived to fit
        '   these ratios.  Across its full range of validity, the correlation fits the reference viscosity ratios to within 0.022 %.

        '   Based on the quality of the fits and the claimed uncertainy of the NIST data,
        '   the uncertainty of viscosities calculated from this correlation is assessed to be 0.3 % (at k = 2).

        '   Inputs
        '       T   -   absolute temperature, in K
        '       P   -   absolute pressure, in Pa

        '   Output
        '       EtaCH4   -   hydrogen gas viscosity, in Pa.s 

        '   Correlation developed by Norman F Glen, August 2020

        '   Code
        '       Version 1.00
        '       Date    20 August 2020
        '       By      Norman Glen

        Try
            '   Correlation limits.
            Const Tmin As Double = 245          ' -28.15°C
            Const Tmax As Double = 330          ' 56.85°C
            Const Pmin As Double = 95000.0#      ' 0.95 bar(a)
            Const Pmax = 2000000.0#               ' 20 bar(a)

            '   Correlation coefficients.
            Dim a(3) As Double
            Dim b(3) As Double

            a(1) = 0.966183
            a(2) = 0.636767
            a(3) = -0.0173905

            b(1) = 1.0#
            b(2) = 0.00124251
            b(3) = 0.00000864591

            '   Reducing parameters.
            Const Tred As Double = 285      ' In K, mid point of temperature range of correlation.

            Dim Tr As Double
            Dim RhoRat As Double
            Dim EtaAmb As Double, EtaRat As Double

            If (T >= Tmin) And (T <= Tmax) And (P >= Pmin) And (P <= Pmax) Then

                Tr = Tred / T

                RhoRat = MethaneProperties(T, P) / MethaneProperties(T, 100000.0#) - 1.0#

                EtaAmb = System.Math.Sqrt(T) / (a(1) + a(2) * Tr + a(3) * Tr ^ 2)
                EtaRat = b(1) + b(2) * RhoRat * Tr ^ 1.3 + b(3) * RhoRat ^ 2 * Tr ^ 2.6

                MethaneViscosity = EtaAmb * EtaRat

                MethaneViscosity /= 1000000  'Convert from micro Pa.s to Pa.s

                Return MethaneViscosity
            Else

                MethaneViscosity = Double.NaN

            End If
        Catch ex As Exception
            ' Log.Error(ex)
            Return Double.NaN
        End Try


    End Function
    Public Function CStarMethane(UpstreamTemperatureABS As Double, UpstreamPressureABS As Double) As Double

        '   Function to calculate C*, critical flow factor, for gaseous methane.

        '   The NIST REFPROP package (NIST Standard Reference Database 23, Version 9.0) was used
        '   to generate the reference values for the development of the correlation.  Values at
        '   atmospheric pressure at 18 temperatures were fitted to a simple cubic against a reduced temperature.
        '   The pressure dependence was derived from C* ratios for 503 values calculated from the REFPROP package.

        '   This implementation uses coefficients with 8 significant digits.

        '   The correlation is valid from 1 to 20 bar(a) between 245 and 330 K.

        '   Although the REFPROP package does not give a references for its C* calculation
        '   the values generated by it agree to 6 signifcant digits with values given in
        '   ISO 9300:2005, which have a stated uncertainty of 0.1 % (k=2).

        '   The correlation immplemented in this function fits the NIST reference data to
        '   within 0.003 %, so the certainty of C* values calculated from the function
        '   at ambient pressure will also be 0.1 % (at k=2).

        '   Across its full range of validity the correlation fits the NIST reference data
        '   to within 0.005 % except at a 19 of the 864 points, 8 of which have errors of up
        '   to -0.012 %.  Examination of plots of C* against pressure in the range from
        '   10,000 Pa to 12,000 Pa show non-smooth variations of C*, probably due to
        '   numerical differentiation artefacts in the reference calculations for small
        '   pressure steps at low pressures.

        '   Based on the quality of the fits and the claimed uncertainy of the ISO 9300:2005 data,
        '   the uncertainty of C* values calculated from this correlation is assessed to be 0.1 % (at k = 2).

        '   Inputs
        '       T   -   absolute temperature, in K
        '       P   -   absolute pressure, in Pa

        '   Output
        '       CStarCH4   -   critical flow factor

        '   Correlation developed by Norman Glen, November 2020.

        '   Code
        '       Version 1.00
        '       Date    03 November 2020
        '       By      Norman Glen

        Try
            '   Correlation limits.
            Dim Tmin As Double = 245          ' -28.15°C
            Dim Tmax As Double = 330          ' 56.85°C
            Dim Pmin As Double = 100000.0#      ' 1 bar(a)
            Dim Pmax = 2000000.0#               ' 20 bar(a)

            '   Correlation coefficients.
            Dim a(3) As Double
            Dim b(4) As Double
            Dim n(2) As Double

            a(0) = 0.576677
            a(1) = 0.212012
            a(2) = -0.158872
            a(3) = 0.0410878

            b(1) = 0.00000638663
            b(2) = 0.0010355
            b(3) = 0.0000186284
            b(4) = -0.0000174943

            n(1) = 3.0#
            n(2) = -0.5

            '   Reducing parameters.
            Const Tred As Double = 285      ' In K, mid point of temperature range of correlation.

            Dim Tr As Double
            Dim RhoRat As Double
            Dim CStarAmb As Double, CStarRat As Double

            Dim i As Integer

            If (UpstreamTemperatureABS >= Tmin) And (UpstreamTemperatureABS <= Tmax) And (UpstreamPressureABS >= Pmin) And (UpstreamPressureABS <= Pmax) Then

                Tr = Tred / UpstreamTemperatureABS

                RhoRat = MethaneProperties(UpstreamTemperatureABS, UpstreamPressureABS) / MethaneProperties(UpstreamTemperatureABS, 100000.0#) - 1.0#

                For i = 0 To 3
                    CStarAmb = CStarAmb + a(i) * Tr ^ i
                Next i

                CStarRat = 1.0# + (b(1) + b(2) * Tr ^ n(1)) * RhoRat + (b(3) + b(4) * Tr ^ n(2)) * RhoRat ^ 2

                CStarMethane = CStarAmb * CStarRat

            Else

                CStarMethane = Double.NaN

            End If
        Catch ex As Exception
            ' Log.Error(ex)
            Return Double.NaN
        End Try


    End Function
    Public Function HydrogenProperties(T As Double, P As Double) As Double

        '   Function to calculate the density or compressibility of gaseous hydrogen.

        '   For each isotherm, the reference compression factor Z, calculated using the NIST REFPROP package
        '   (NIST Standard Reference Database 23, Version 9.0) implementation of the reference-quality formulation
        '   of  Leachman, J.W., Jacobsen, R.T, Penoncello, S.G., Lemmon, E.W.
        '   "Fundamental Equations of State for Parahydrogen, Normal Hydrogen, and Orthohydrogen",
        '   J. Phys. Chem. Ref. Data, 38(3):721-748, 2009 was fitted to a simplified NEL-style equation
        '   in normalised temperature and normalised pressure.

        '   This implementation uses coefficients with 6 significant digits.

        '   The correlation is valid from 0.95 to 20 bar(a) between 245 and 330 K.  Across its full range of validity,
        '   the correlation fits the reference data to within 0.0006 %.  The uncertainty of the reference data is 0.04 % (at k = 2),
        '   so the uncertainty of densities calculated from this correlation will also be 0.04 % (at k = 2).

        '   Inputs
        '       T   -   absolute temperature, in K
        '       P   -   absolute pressure, in Pa

        '   Output
        '       RhoH2   -   hydrogen gas density, in kg/m3

        '   Correlation developed by Norman F Glen, February 2019.

        '   Code
        '       Version 1.00
        '       Date    06 February 2019
        '       By      Norman Glen

        '       Version 1.10
        '       Date    23 June 2020
        '       By      Norman Glen
        '       Notes   Temperature range extended down from 270 K to 245 K and up from 300 K to 330 K

        '       Version 1.20
        '       Date    07 August 2020
        '       By      Norman Glen
        '       Notes   Correlation coefficients for compressibility now optimised against reference density data.


        Try
            '   Correlation limits.
            Dim Tmin As Double = 245          ' -28.15°C
            Dim Tmax As Double = 330          ' 56.85°C
            Dim Pmin As Double = 95000.0#      ' 0.95 bar(a)
            Dim Pmax = 2000000.0#               ' 20 bar(a)

            '   Correlation coefficients.
            Dim a(3) As Double

            a(1) = 0.00737858
            a(2) = -0.00136445
            a(3) = 0.0000223767

            '   Reducing parameters.
            Const Tred As Double = 285      ' In K, mid point of temperature range of correlation.
            Const Pred As Double = 1000000.0# ' In Pa, mid point of temperature range of correlation.

            '   Universal gas constant.
            Const R As Double = 8.31446261815824    ' In J/(K.mol) - exact numerical value following re-definition of SI base units, 20 May 2019.

            '   Molar mass of hydrogen.
            Const MH2 As Double = 2.015882          ' In g/mol - from NIST database (January 2015) for molar mass and isotopic composition for hydrogen.

            '   Specific gas constant.
            Dim Rspec As Double
            Rspec = 1000 * R / MH2                  ' In J/(kg.K).

            Dim Tr As Double, Pr As Double
            Dim Z As Double
            Dim RhoH2 As Double

            Tr = Tred / T
            Pr = P / Pred

            Z = 1 + Pr * (a(1) * Tr + a(2) * Tr ^ 3 + a(3) * Pr ^ 0.75 * Tr ^ 5)

            RhoH2 = P / (Rspec * T * Z)

            Return RhoH2

        Catch ex As Exception
            ' Log.Error(ex)
            Return Double.NaN
        End Try


    End Function
    Public Function HydrogenViscosity(T As Double, P As Double) As Double

        '   Function to calculate the viscosity of gaseous hydrogen.

        '   A simplified form of the NEL equation developed by Watson for hydrogen was used to fit viscosity at ambient pressure
        '   ('Viscosity of Gaaes in Metric Units', National Engineering Laboratory, HMSO, 1972).  The pressure dependance was
        '   derived from viscosity ratios for 864 values calculated using the NIST REFPROP package
        '   (NIST Standard Reference Database 23, Version 9.0) implementation of the reference-quality formulation
        '   of Leachman, J.W., Jacobsen, R.T, Penoncello, S.G., Lemmon, E.W.
        '   "Fundamental Equations of State for Parahydrogen, Normal Hydrogen, and Orthohydrogen",
        '   J. Phys. Chem. Ref. Data, 38(3):721-748, 2009.

        '   This implementation uses coefficients with 6 significant digits.

        '   The correlation is valid from 0.95 to 20 bar(a) between 245 and 330 K.

        '   The uncertainty in the NEL reference data at ambient pressure is 1.0 % (at k=2) and the correlation implemented
        '   in this function fits these data to within 0.01 %, so the certainty of viscosities calculated from the function
        '   at ambient pressure will also be 1.0 % (at k=2).  Although only given at ambient pressure, the NEL reference data
        '   are claimed to be valid to within 1.0 % at up to 80 bar below 300 K.

        '   The uncertainty of the NIST reference data is between 4 and 15 % (at k=2).  These data were therefore only used to determmine
        '   the viscosity ratio at each isotherm and a function derived to fit these ratios.  Across its full range of validity,
        '   the correlation fits the reference viscosity ratios to within 0.01 %.

        '   Based on the quality of the fits and the claimed uncertainy of the NEL ambient pressure data,
        '   the uncertainty of viscosities calculated from this correlation is assessed to be 1.0 % (at k = 2).

        '   Inputs
        '       T   -   absolute temperature, in K
        '       P   -   absolute pressure, in Pa

        '   Output
        '       EtaH2   -   hydrogen gas viscosity, in Pa.s

        '   Correlation developed by Norman F Glen, February 2019.

        '   Code
        '       Version 1.00
        '       Date    14 February 2019
        '       By      Norman Glen

        '       Version 1.10
        '       Date    23 June 2020
        '       By      Norman Glen
        '       Notes   Temperature range extended down from 270 K to 245 K and up from 300 K to 330 K

        Try
            '   Correlation limits.
            Dim Tmin As Double = 245          ' -28.15°C
            Dim Tmax As Double = 330          ' 56.85°C
            Dim Pmin As Double = 95000.0#      ' 0.95 bar(a)
            Dim Pmax = 2000000.0#               ' 20 bar(a)

            '   Correlation coefficients.
            Dim a(3) As Double
            Dim b(3) As Double

            a(1) = 1.80511
            a(2) = 0.143897

            b(1) = 0.999708
            b(2) = -0.0000857637
            b(3) = 0.00299197

            '   Reducing parameters.
            Const Tred As Double = 285      ' In K, mid point of temperature range of correlation.
            Const Pred As Double = 1000000.0# ' In Pa, mid point of temperature range of correlation.

            Dim Tr As Double, Pr As Double
            Dim EtaAmb As Double, EtaRat As Double

            Tr = Tred / T
            Pr = P / Pred

            EtaAmb = System.Math.Sqrt(T) / (a(1) + a(2) * Tr ^ 3)
            EtaRat = b(1) + Pr * (b(2) + b(3) * Tr ^ 1.5)

            HydrogenViscosity = EtaAmb * EtaRat

            HydrogenViscosity /= 1000000  'Convert from micro Pa.s to Pa.s

            Return HydrogenViscosity

        Catch ex As Exception
            ' Log.Error(ex)
            Return Double.NaN
        End Try

    End Function
    Public Function CStarHydrogen(UpstreamTemperatureABS As Double, UpstreamPressureABS As Double) As Double

        '   Function to calculate C*, critical flow factor, for gaseous hydrogen.

        '   This implementation uses coefficients with 8 significant digits.

        '   The correlation is valid from 0.95 to 20 bar(a) between 245 and 330 K.

        '   Inputs
        '       UpstreamTemperatureABS   -   absolute temperature, in K
        '       UpstreamPressureABS   -   absolute pressure, in Pa

        '   Output
        '       CStar   -   critical flow factor

        '   Correlation developed by Gary Corpron, CEESI, based on data from Robert Johnston, NASA.

        '   Code
        '       Version 1.00
        '       Date    14 February 2019
        '       By      Norman Glen

        '       Version 1.10
        '       Date    23 June 2020
        '       By      Norman Glen
        '       Notes   Temperature range extended down from 270 K to 245 K and up from 300 K to 330 K

        Try

            '   Correlation limits.
            Dim Tmin As Double = 245          ' -28.15°C
            Dim Tmax As Double = 330          ' 56.85°C
            Dim Pmin As Double = 95000.0#      ' 0.95 bar(a)
            Dim Pmax = 2000000.0#               ' 20 bar(a)

            '   Correlation coefficients.
            Dim a(6) As Double

            a(0) = 0.79741185
            a(1) = -0.33912011
            a(2) = 0.00029854078
            a(3) = 0.33862248
            a(4) = -0.0010015041
            a(5) = -0.11242827
            a(6) = 0.00067411915

            '   Reducing parameters.
            Const Tred As Double = 100
            Const Pred As Double = 100000.0# ' To convert input in Pa to bar, as correlation was developed in bar.

            Dim Tr As Double, Pr As Double

            Tr = System.Math.Log10(1.8 * UpstreamTemperatureABS / Tred)
            Pr = UpstreamPressureABS / Pred

            CStarHydrogen = a(0) + a(1) * Tr + a(2) * Pr + a(3) * Tr ^ 2 + a(4) * Pr * Tr + a(5) * Tr ^ 3 + a(6) * Pr * Tr ^ 2

        Catch ex As Exception
            ' Log.Error(ex)
            Return Double.NaN
        End Try

    End Function
End Class