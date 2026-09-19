import React, { useState, useEffect, useCallback } from "react";
import {
  LayoutDashboard, Users, Building2, Clock, Plus, Pencil, Trash2,
  ChevronLeft, ChevronRight, LogOut, X, Lock, Unlock, AlertTriangle,
  Check, Search, CalendarDays, CalendarRange, ClipboardCheck, Euro, Home, SlidersHorizontal
} from "lucide-react";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell, Legend
} from "recharts";
import * as XLSX from "xlsx";
import { supabase, supabaseConfigured } from "./supabaseClient";

/* ---------------------------------------------------------------------- */
/* Design tokens                                                          */
/* ---------------------------------------------------------------------- */

const C = {
  bg: "#FAFAF8",
  surface: "#FFFFFF",
  ink: "#1B1E1C",
  inkMuted: "#6E7570",
  inkFaint: "#9AA19B",
  border: "#E3E1DA",
  borderStrong: "#CFCCC2",
  sidebarBg: "#12211E",
  sidebarBgActive: "#1E3A34",
  sidebarText: "#B9CCC5",
  sidebarTextActive: "#FFFFFF",
  accent: "#1C7C67",
  accentDark: "#145F4F",
  accentSoft: "#E6F1EE",
  warn: "#A86A17",
  warnSoft: "#FBF0DD",
  danger: "#AE362B",
  dangerSoft: "#FBEAE8",
};

const sans =
  "'IBM Plex Sans', ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, sans-serif";
const mono =
  "'IBM Plex Mono', ui-monospace, 'SF Mono', 'Roboto Mono', Menlo, monospace";

const LOGO_DATA_URI = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMgAAADICAYAAACtWK6eAAAABHNCSVQICAgIfAhkiAAAHUNJREFUeJztnXuUHFWdx7+/W9XzzuQByZCQEAgkgUAg4RGSsBjBFXVddeMaVoSQQBR0d3VFTFTUjdldRUAeHs8eFxHYdc+e1bhHlJiI+OAICfJGeYWQ50xIhskkM5lXTz+q7m//6AxMeqqqq7ur6lZ33885c3JSfetX3+mpb9V9/i6g0Wg0Go1Go9FoNBqNRqPRaDQajUaj0Wg0Go3mHUi1gKBZ+etFn4fERap1AIAQRvupT/7xaxs2QKrWoikNoVpAkFy3ZdG8uJgDAKS0Z+5bvOQy1To0pVM1Blm/HoINrFStIx9bWitu2HRBk2odmtKoGoPsXbL4clviFNU68iFB47ImfVS1Dk1pVIVBVj56brNk+2OqdbhhkXHF9b+5ZJpqHZriqQqDQNZ9lEDjVMtwhdmwbfsa1TI0xVPxBrn2l4tPBoz3qtZRCMnyvFWbly5QrUNTHBVvEDJxNZgN1Tr8IMm6ZsXGFRWhVZOjog2y6tGlCyTL81Tr8A1havOEjveplqHxT8UaZMXGFQZnZey6dQshJS+/esuiVtU6NP6oWIM0T+h4Hwt5kmodxcKMJkOIK1Xr0PijIg1y9ZZFrVLyctU6SkWyXLZm85KZqnVoClORBjGEuJIZFTs6TYBIG/Ja1To0hak4g6zZvGSmZLlMtY5yIeYzV21ZtFi1Do03FWeQtGmtpArU7YQkXLVi47w61To07lTUjbZqy6LFJOks1ToC5MTGcc0fVC1C407FGGTFxnl1knCVah1BIyE+tOaRxZNU69A4UzEGOfakPVG1jqAhcL0Fq+qMXy1UhEHWPLJ4koT4kGodYSFJLLl280VzVOvQjKUiDJIi++MErletIzQkiAWtBFffEuhKJ/YGWb3p4tkCtFS1jgiYtfJXi9+lWoTmeOJtEAZlBV8LWRtPViJ55arHljWo1qF5h1gb5NotSy4VArNU64gKBiZwerBip9BUI7E1yKrHljVA2H+nWkfUkDTef9XDF7ep1qHJEVuDIJX6GwYmqJYRNZJg1tfJq1Xr0OSIpUFWPrp0CoPfr1qHKmymC1ZvWTxftQ5NTA3CtnU1AwnVOlRiC6mX58aA2Blk1SMXnU3Ahap1KIcxvallz3tUy6h1YmWQ9eshWFLFLaMNCwnjYys2LmtRraOWiZVB2hdd/B4mzFCtIzYQmhvGJWObEK8WiJVBbENXrcYgxIWsp6AoIz5f/K08a0rjngeXnvx/p6qWAgDNjU/8ialL+WKmvkz64Rd7/nIhsRGLLm8m62cH19z9Y9U6oiJGBpE/A9HyRVP+iLbmvarVwDAOdjXW//h1lRpsll3PHprbkRUzvqhSx2iY+QixPefNT97do1pLFMSjinUrXw6i5QDwas8CSDZVK4JtT22z7DPHq9RwcCC7LUMnr1apIR8iOkHC2KBaR1SoN8gKNgC+Z+S/Q1Yj9h49W6WiYxBS6aVnMAwlb9m0bb3ePnzRxUQifovECDe2/eBLcfgjhY56g5yPG0F03Kjxzr4zkbGaVSkaxYSWjLUo8uR0DFi7+htes2lSLHuwBFHCENbdqnVEgVqD3MoTAR7zus6ygdd7zlehaAyZ9PmnAeMiHdEetLJP92Yu/FtBQnkngRtCiPe2PfjFql3lOYJagzB/A0SOVYiOoRnoS6mf1EpUn0hlLo0sC6IE92/vmTEkqeHSqK5ZKqbEnfM2ro+tiYNAnUG+xfNA/Bm3jxnAaz0XIA4dbdnsnOk2T22M4lq9Q+kn0jz7OiKh/hcvABHNPto3+E+qdYSJOoMQ3w0SnhMSD6cn4ODAGVEpcoVIUDp9+elhXycr5Zs7hs6dAZFQ/0v7xaCvTbnvn9S/6kNCjUG+zR8C0RV+im7vPRdSqn+LS9l2QtaeH1r+KibizkE8lcXUipqLJkCtJpnfUq0jLKI3yL2cAPOdfosn7XrsOnpOmIp8k8ksOZ1FXShVn1Q2/UrH8AWXCYrHiHkxENGqqffdfIFqHWEQvUF65OdBNLuYU3b3z0Eqq37PGeaWpkx6cfC71TJndh6dsFuKCR8OPHYEEJFBgu5GHBqMAROtQb7JbWB8rdjTLBbY3hOPB1Qme96pjPGBLubqz6af7LMXXEUg9VMISkQQXTrtgbVVtzFQtAYh/iaEKOlV8GZyKnqH1W81TkiYqfS7Twsqng3Zu6N/NkM0LAoqpiqI5R3T77opkt6+qIjOILfz+SCsLifEqz3ng1n94L9lzzrJkqcEMtTfnbQeH5Yzrw8ilmqIxAxMNNap1hEkEd1tTJB8N4jKGpHuzbTiwMDcoESVDIEom3l32V2xGbb37h5YOJeo7pQgdMUBlrxu2gNfqJpFb9EY5NtYAVAgaTVfO3oOLKk++aAtT5yQtRZOLvV8BuT+AfN5myZ/IkhdqiESTQDdplpHUIRvkLu4Ecx3BBUubSews/fcoMKVRTqzeBZTQ0nf4XA286fO1PnvJxLjgtalGmJ8fPIP116iWkcQhG+QFNaCKNAqxJ7+0zGUnRhkyBJpbEhnLple7FkMDO/on7JfovUDYahSDZGgBPE9qIK0ReEa5DaeDuIvBR1WgrD9SDy6fa3M2adInlTU1gxH08mtg/a8VUSkvschJATRhW2DMyt+J99w/0A23waiULZr7hyegsPJGLRtyTTSmct9d/tazN2v953VBGo4L0xZccBg+tYJ96+r6CpkeAb5Ni9FyHsKvtKzEMzq3+K2PaPNkqf7Gt/pGsg8kaWZ14WtKQ4Iwkn1xEUPDMeJcAyyngWk/C6IQp16MJBtRkd/HDa9JaTT7zqj0NeZkfYbHakLFxLMqREJUw/z52f859rQZ0KHRTgGacC1ECKSHFc7js5D1g6lFlcUzBPHZawLXad9M2DtOVr3UhaTqm46hheCqM625V2qdZRK8Aa5jccBfGvgcV1ISxNv9C6I6nKeZDIXzZJocqzzDVmZ57rtCz5CZKh3c8QIEh+eet+696rWUQrBG0TiqyCKNNHB3oFTMZCOQ/KP+rpMetmYngMbGNzVO62Hualmk1GTkHfhsfUVNxkzWIPczqcD/PlAY/rgneW56slac6ZL2XbcUH9fKvV4v5y7uhKW0YaFIDpn2s7+T6vWUSzBGsTi74BIyXbNh1InoGtI/XaGRIZIZS57u1GalbLzjb75k0nUx6E3QSlkGN+Y/sObQluVGQbBGeQ2fg8E/U1g8Urg1Z7zYpGVUcppJ2bteRMBoDNpbY1bdkRVENEJNpv/olpHMQTzyt/IBnbzi/kJ4FRwaksHJtR3q5YBxsDA9PHr7n5pcP5sxgmhjgdVEgy2LMtY0HXDba+q1uKHYB63e+3lIEO5OQBg3+ApwGAMRtiBcXv6/31qQ/NLHxXhDgf5gigD09yjWgYIZBqG9RUA16jW4oeA6iPGr8HcBaKqTf9SLATrqODJdenhJV3CyMTCsY2NwzCMTtUyYED8SLUGvwT3aLuVrwfh/sDiVThNRtcmg5Bl2A2G2X0FCfWNIyH60dS0Gbl+PzVIKZ/JGuJGrzLjB1Lbd33ue+moNHkR3B/tBfwXFvJnIKjmd4kyaHhvgrP9khKNBCPFsnEHiaTybOhStiKbnYNEYoeS67NkycD2hC1Xe5VLNtX/DsCmaFR5E1wv1k/JhqCbwKzu8RQLmBuNnqekSLydvIDt1l1SiiGVqkbIZOaDWUlPPBh4kYCjhcqRwcum3ntzHEZ+Ax4H+TJtBeMngcasMBKi/wVD5r2ZiWzYLS8rknQczHXIZKNfkckskyToGV9lbZgw5UfC1uSH4Kea1NM6MA8HHrcSYJluFL2vSGGOmY/F3HxQSlN9/zOAbOYMSBltAkcmepKlzPgtTyzmn/iDL8wJU5MfgjfIzbQfonoW7RdDg+jdymh0XSDE9viXOBZVUEI6Hd3UHMl8CMBrxZ5Xb4jlWL9e6arLcC6exB2QvD+U2DGFkO2uNwbe9CzEdX1SNqjfoRQAuAnM0WzByMSPQxb/YGDmqVNPGVwahia/hGOQDZSESWtDiR1TmozDW6XH22MElhNeY4lsFJo8VMAw34JtT0PY6XRZ8hskcaDU8wn4wNR7b1C2RCC819c6bATzE6HFjxEGkm8ISF/tLmKRYdlSdHUjSAzjCIgyYK6DlCWn9iqIBFsM3lpWEEYzm+OVZX8JsX5HDNBNYLbDu0YMYJbNxpFnQKbvbHZsteyVthgIU5YbRBZM8/Db/7ftNgQ5HJbH8wDK/j0N8NKZD66NfDNVIOysJl+h5yHxX6FeQzF1Zv/TTHXFpXoUJCHHvRSSJE9M4xCA0c8sAdsOfom8lHKQmZ4LIhYzG5KxPIhYxRJBDwHdAub+8K+jAnuoSfTtQAnfI3NTl5SJSCdGCRqGMHrHHJdyEpiDruaLrcTSCiqaLeXctvu/GPlOSuEb5KvUBfC/hX4dBTTQ4a0215e8sw9nW19mhgxSkxeG+ZbrZ7kGezBI5oOADHw+Sx3hI9i4MdI8T9H0MafEd8G8M5JrRYRAprNepLrKCkJ1gywbdwckyRND9EGIpOvnzM2BDB4eG+d5vOxADtiSJ5889MyyMGK7EY1BNlAGRDdHcq2IaEp0b5XU0FJuHGm1vs6SQp65KmGYhb0s7Wko95Zg4u1gLu/B4RWfccXcCLM1RjdK+WXaBMm/iex6IWJS8hWSMpDeOSKRlTwu1NV1pnkERIWHXhgJSDml5OtI5gyYtpUcwN9FGgZE9oOhXmMU0Q7jC7oJYMWDZGXC0moS3S+A6oLbm9pqbIcUBWe5lgJRFoZxuHDBY9j2ZAAl/mpMz4LZvR4XEMRiUVSb9ERrkC/Tq2DcG+k1A6ZOHH2S0RDI9mtvQ4JtK5xuX9PsAorqByit21cyHwX4xaJPLAGWEAwzkm7f6CeC1dN6MB+J/LoBQLD7mszBPaAw5mc0HZZ2ouQpGU4IkYQQfUWfJ+UEMBf3DGDirUB0g8Ik7Vkzf/iFhWFfJ3qDfIF6IGl95NcNgAbjSFnduoVga8LLzAjsJjM9unULYdsn+y4rJe8niUh640ZjG+LDF9x7Q6BbcuejZirxbPwHJL+i5NolIpDuqMNwuG8+MpNsNwXSHW4YvSAqfVkOcyOkLJzjjZklBP5Q8oXKQNo88YAxPtR0rmoMciXZILpJybVLpNns2hZEt24h2G7dIVmUueBMwjAOla0l1xYpNC7HL0OqqzITyctnPrg+tNVf6hajfIV+C8hfKLt+EZjofxGUiOa7IrLZai7r7Wqa3SAKYpaHCdt27/aVzClmeiqAC5UMMeosmfxwWPHV7pEnxM1g9r0MUwlsZ5rMnpfAItS67vHXbNnP0izpqUyUgWEE90CXcjLgkuRBMJ4COBXYxUqEwQvb7vui723wikGtQdbRbjDuUaqhAA1GzxNey2jDQtqtJXX75hrmQa7qJdhy7DwtZj4iiWKRiALMVCewHMyB9y6q32XVoH8Dc+ndLSFCyPQ0GMNqlg5zfa+069uLOUWIQYgQlplI2Qrm458RLPE4WEY20bIQNvMpbQ+uXRR0XPUG+RINQPAtqmU40SR6Hg+zW7cQ0h7/Kkv4bEzwsUHBcBi9PFdC7gFxR2gXKxFh44NnbPlsoEm/1BsEAJ4TP4KUgSyuCQrByV0mZQdVashlZWzyNW08160bXnOAuQFSngBmtlkilkupibg12VV/RaAxgwxWFt/iS0D8M9Uy3iH7/RajZ17hbs5wISENJA5dRkC9+5QRSfV1OxIU3BijixgbZmL7i+Dsk+FeqHTIgCUz+HbnjXf6n4DmgfKEym9zC20DEJ/s8LfyBYNWW7BzrkolM8Vz34LWlmfnAObJHPKfkzidBFkvKsx9XZBRWRkDSaQejypWHPkKPQ/EP7dXoq67pb5uf3DLAT2gxIE9peS3ipogszJqg3ghMj9XLaEQ0m60WFLoNy3RYI+BQ6FMyQ8cJq4zEUj6W20QL77c2AFBz6uW4YVtt6SGs6eF3KPEbBjt8cgI6QM2+JmD198VyNtfG6QQw9gCgViP9g8lz+yQ3BDasl1hHDlINFAZCckFpVrZ+GVg4YIKVLVsoH7Y8veqZXjCDXZyeG5IGxDKrCHaYzfm4QaR+eiONbcHNlqqDeKHduNxMPeoluHFcPqMt2y7NfD8Y4ZxYB+QrYzsmEIcPtB8fqBT77VB/PADygJis2oZhRhIzg80tRIhNSjoYPnz5iPCZv45rrwyUDNrg/jlq/QSJEW+aq4YstZJ/Zns1MDmtQmzfQ8o/t26AGAIsaNrzXcCX4SnDVIMhF+A4zxMBvQn5+9hNst+ipIY6BbUWxEpY0lACsJDYcTWBimGr1InhHxatQwvWI5LpzKnFjULeAwEmTD37gtGUfjYLLa1X3dHKDPCtUGKxTQeAQUzCBUWg8mz90vZVLJGIbrfZE7GYp/yghCGyGr6VVjhtUGKZS0NwbZ/q1qGN6YcHJ5bUnuJyEobRrv3VnIxgmH+qvPGDaElq9MGKYU/m9sgOBY71rqRzszqzsqJY/c6KIAwO/eBrdgshPKCDers7GgIdWaxNkgp/JRsSPGwahmFGBpauBPkf54WUapf4ECsjT+arEg8hA0bQjWzNkip3EKvw+bXVcvwImtNHEpnZhz0VZiIhbEvpNH44GGSLx9edesbYV9HG6QcGsQmhL5KqTwGUufsZSQKLtsl9B4SdFTpCkq/kAELlogkZZQ2SDmspUNgGdvVdQDAVmM2mZldYCYu20aivbyu4Qhhm/4Q1IrBQmiDlEvaeBTEQ6pleJFMzj1gyxZXjcI4tJ94ONYzlkdgpv6maelHo7qeNki5bKAUhHhEtQxPWPBgat4up48ImZRhdgSaVT5MjDrxy11/9b3Ixmi0QYJgHZ4GONIda4slk57RY9snjqmWGEbnHkg71tNnRjCIOt689vZno7ymNkgQEDHsbOzzDPclF+4CRnX70tBREp2xnsb/NkSckXgIFO1cOG2QoPh6w24YFMouUUFhW63DqczMY0tRiQ1zX8UsozWZX+j61Hci16sNEiRJbAYh1nswDqbOaZdoyJA48pbAQKw7F0ZgQoYMa5OKa2uDBMkG6gEolD3Cg4Lteis5POt1Q+yvmG5dEH7Xft13lWRU0QYJmmY8BlDxGwNGyHBmzgNkpCuieiUM6p2WHVCWE0AbJGg+R2kQQpt+XTZEb2HYfNJE9qFi5mmpIg36xfM3/kBZtVUbJAzinJXRwMPYQLLj+nv2ECiSbZtLRRB2d193x5+UalB58aomjlkZiV/FOno7sYMpmh5mUCw7FUhAssyEsoy2GLRBwiJuWRkJNlgc1xPUft2Gowz6nSpJXjDE0wc+9T3lC7fik929GmnCQxigx1TLAACkkc31sh3PyXbf77rqWi+WNk9UIcsREsMThps2+5unHy7x2R9Eo4ypD6w9n6S8VrWOEWwyftG15vZYPFi0QTQAgMkPrj1JtYYRuvc1HQp7paBGo9FoNBqNRqPRaDQajUaj0Wg0Go1Go9FoNBqNRqPRaDQaTSyo1dm8J8L9d+8Bys7YPgnAOQDaALQA6AXQCeAFoKy0QHUA5gOYAWACgCEAhwH8GTnd5TANwDzktLcA6Aew/1jsVJmxNRXEeAASALv8vKuM2FcAeAw5gznF7gPwPwDmFhn3NAAPInfTOsW1AGwD8KEi4yYAfBrAKy5xGUASwE8AnFdkbE2F8i643wwM4HMlxEwAuL9A3NE/GQCrfcZeCWCwiNg/Qe5NU4hTATxbRNwsgFt8atZUMJ+F943wwyLjEYD/LRDT6UcC+KsCsa+C+9vI6+f+AnFPArCnhLgM4F8LxNZUOPfB+wZ4rsh4qwvE8/rZD/en/Qy4V6n8mO8SD80PlaHZArDEx/eiqVCegfcNMAz/ySwSAA66xDkI4OsArgHwY4/r/a1L7AdcylsAvo9c1etfjul1Kvcjl7jLPLS0A/gugA0Afg73t9dj/r4eTaVhINfoLPSUPNtnvE+4nN+PXMN6ND91Kft9h7iTPHR+Kq/sSpdyb7lo/plL+RcAjMsre7VLWUauN01TZZwFf9WIq33Gc7vp73Uoe5FL2W0OZa9xKfsmciYfDQHY51J+Sl7ZZri/cd7v8js+7VL+6y7lq4paSxznt6vSTzkDwF+6fLbF4djzyHXz5nOiw7H3usT9PcaO0TAAt+RvJ+T9/zIADQ7l+gH8xiWG28ZAboaqKrRBnBvlfgxyCnKDdU684nBMAtjhcDxRxPVfcznuFBcYawa3uNvhPjj6Z5fjNVHF0gYB/ttnuXxme3zW5XLcaR/D/rz/k0dst7huWyIP5P1/jks5r0083ZJwjwMw1eO8qqDWDLLA4dgTyDWIR9OGwn/8GR6fDbscd9pqOb/aNQlAk8v5gy7H8/WPkG++6S7l8o3kJzYAnOzxWVVQSwaZDOebficApy2Szy0Qr8Xjs2Lmcu0uIq6b8Zyu1w+g22dsr7lWlsdnbkauGmrJIE7Vpk7knspOBnF624ymuWxFOfLbD143XTHpOLcj14D3E9vLBF54mbkqqKXs7k43/O68f0dTqB3i1LguxCcB/EPesfzqTSk33cMY26XrVJ2r+hs6aGrJIE5Vpl15/44mjNmr+W2CoEhjbHVKEwC1VMVyeoN4GWQugMbw5GgqgVoxSB2AMx2Oj2xH5mQQA7lFT5oaplYMchac2wwjBnkTzj05epFQjVMrBnG70Uca5xLA3iLO09QItWyQQzi+0bzToUyhrl5NlVPLBslvdzi1Q+ajdhNbaKANMhqnsZDxGLuuQ1ND1IJBpsN5Snl+lcrpDQLodkhNUwsGcZtT5aeKBWiD1DS1YJBCPVgjtMN5eoZuqNcwtWAQtxs8/41hQ3f1avKoBYM4VbEOI5cONJ89Dsdmwn3loKbKqfbJis1wXp3nNOYxcvwDeccIOZM9nnc8fyq5HxYCOMPh+MPITTgEnKt5hZgM4N0Ox59FLqHDCKXErmmq3SBnY2wWECC38GiZw3G372MBxhpkqAQ9azB2ujuQm6o+Mhu3lLjzAGx0OL4axxuklNhepAsXqWyq3SBu7Y/Lj/34xamaFvTN5ieuk9mDiF1qVTus7yA2VHsbpNCyWb84Ge2IR/lybuSjcF856JSypxjctkjwmtbv9bt4fQdVQbUbJKgeqLMx9m3rNm4CFL9We/Sa8mG4ZxIpZRXjaN5wOd7qcY7b75KBc69fVVHNBhlpXAdBA8bu6bET7k96p5F7AKh3OZ6/0tAtz1X+stoR3G7w/CwobnFPcjkOuGdC2YHS17JXDNXcBjkNzjfOQQDf9DhvFYBFDscXAHh11P8HAbwE5+rXbDg/XZ0WbfVj7I22DbnNePI5y+EY4L4hT35XtlOaUyC3MMyAc3YUt3bc712OayqE5XDOKeuUFnQ0/+py3u0OZe90KXuXQ9lTkTOCn9y8i13itsO5TfBHl/JOb5x2l7JuqUSfcinvlnZVUyF8A85/2HsKnOeWOPrXDmXPgfN2bkkc/xZqQM6YTnHvdIhLeCdtT/7P+ryyf+9Szq2N9M8u5bcjlzBvNJ9zKfsGyu9R0yjGbZOYfyxwnlsWdrftBNxu/AyA3yI3PuH21Ga457i9weOcPx+Lu82jTL6RRpiMXE+Z0zl9ADYB+E/ktkNwi73aJbamgnDbYsypbj+aVpfzGM6ZGc9AcXsIjv55yEOHAffqTaGfw/BueH+yxLiMnOn1IrIKx2snWz8LoA64nPs+l/IfRO6NUcyNthtjtyfI56Rj5YqJmwHw1z5+R7f2k9fPcwAm+oitiTmXwvkPnIa/uvPvXc7/ksc5fwH/G2M+glxVxw+TAWz2GbcDzlNo3Pg0ctWqQnFt5DYFqvpcvPlUazdvHXK7P+XTBX+JpX8K5y0FvDIjbkWuG/cqAB9Hrm0xBbnBvSRyqYWeQG67hT/40DBCN3JvqEuR235tCXJvlhbkfpcu5Dbn2YTcvoTFzI/6D+TaMp9Ebo/12cc0E3LtlP3IVakegPOeJxqNRqPRaDQajUaj0Wg0Go1Go9FoNBqNRqPRaGqK/weS0/z2VZ1vBAAAAABJRU5ErkJggg==";

const FontImport = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
    * { box-sizing: border-box; }
    ::selection { background: ${C.accentSoft}; }
    input:focus, select:focus, button:focus, textarea:focus {
      outline: 2px solid ${C.accent}; outline-offset: 1px;
    }
    input::placeholder { color: ${C.inkFaint}; }
    .scrollbar-thin::-webkit-scrollbar { height: 6px; width: 6px; }
    .scrollbar-thin::-webkit-scrollbar-thumb { background: ${C.borderStrong}; border-radius: 3px; }
  `}</style>
);

// Midnight today. Every date the app compares is a local midnight (fromKey,
// startOfMonth, ...), so "today" must be one too — otherwise Monday's
// entries (00:00) fall before a week start that still carries the current
// time of day, and silently disappear from the week view.
const TODAY = (() => { const d = new Date(); d.setHours(0, 0, 0, 0); return d; })();
const CATEGORIES = ["accounting", "tax", "payroll", "other"];
const CATEGORY_LABELS = { accounting: "Accounting", tax: "Tax", payroll: "Payroll", other: "Other" };

// Internal, non-chargeable activities that can be logged instead of a
// client. They count toward the employee's hours (weekly target, compliance)
// but carry no revenue and are left out of the client cost attribution and
// of the "by hours" cost-center driver.
const ACTIVITIES = ["business_development", "training", "client_communication", "other_non_chargeable"];
const ACTIVITY_LABELS = { business_development: "Business Development", training: "Training", client_communication: "Communication with Clients", other_non_chargeable: "Other non-chargeable" };
const ACTIVITY_PREFIX = "act:"; // select values: a client id, or "act:<activity>"
const isActivityValue = (v) => typeof v === "string" && v.startsWith(ACTIVITY_PREFIX);

// Cost centers — the reporting dimension applied to every euro of revenue
// (fixed fees split per client, extra fees tagged one by one) and of labor
// cost (per employee: a fixed cost center, or spread by their logged hours).
const COST_CENTERS = ["accounting", "tax", "payroll", "other", "director", "finance", "rental"];
const COST_CENTER_LABELS = { accounting: "Accounting", tax: "Tax", payroll: "Payroll", other: "Other", director: "Director", finance: "Finance", rental: "Rental" };
// Departments group cost centers for the P&L: each cost center belongs to
// exactly one department, so department totals reconcile to the company.
const DEPARTMENTS = [
  { key: "accounting_team", label: "Accounting Team ACCO", costCenters: ["accounting", "tax", "payroll", "other"] },
  { key: "management_team", label: "Management Team", costCenters: ["director", "finance"] },
  { key: "rental", label: "Rental", costCenters: ["rental"] },
];
const departmentOf = (cc) => DEPARTMENTS.find((d) => d.costCenters.includes(cc));

// Rental expenses: booked per property and per month, in three categories.
// They are non-labor cost of the Rental department / Rental cost center.
const RENTAL_EXPENSE_CATEGORIES = ["rent", "utilities", "common_expenses"];
const RENTAL_EXPENSE_LABELS = { rent: "Rental", utilities: "Utilities", common_expenses: "Common expenses" };
// An expense is either recurring (a monthly amount from one month to another,
// or open-ended) or extra (a one-off amount booked to a single month).
const RENTAL_EXPENSE_KINDS = { recurring: "Recurring (from – to)", extra: "Extra (one month)" };
const monthKeyOf = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
const fmtMonthKey = (k) => k ? fromKey(k + "-01").toLocaleDateString("en-GB", { month: "short", year: "numeric" }) : "";
// Does a rental expense apply to a given 'YYYY-MM'?
const rentalExpenseAppliesTo = (x, monthKey) => x.kind === "recurring"
  ? (!!x.fromMonth && x.fromMonth <= monthKey && (!x.toMonth || monthKey <= x.toMonth))
  : x.month === monthKey;
const UNALLOCATED = "unallocated"; // fixed fees with no split yet, or cost of employees with no hours and no fixed cost center
const DEFAULT_DEPARTMENT = "accounting_team";
const departmentLabel = (key) => DEPARTMENTS.find((d) => d.key === key)?.label || key;
const emptyByCC = () => Object.fromEntries([...COST_CENTERS, UNALLOCATED].map((k) => [k, 0]));
const emptyFeeSplit = () => Object.fromEntries(COST_CENTERS.map((k) => [k, 0]));
const feeSplitTotal = (s) => COST_CENTERS.reduce((t, k) => t + (Number(s?.[k]) || 0), 0);
const costCenterLabel = (cc) => cc === UNALLOCATED ? "Unallocated" : (COST_CENTER_LABELS[cc] || cc);

/* ---------------------------------------------------------------------- */
/* Date helpers                                                            */
/* ---------------------------------------------------------------------- */

const pad = (n) => String(n).padStart(2, "0");
const toKey = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const fromKey = (k) => { const [y, m, d] = k.split("-").map(Number); return new Date(y, m - 1, d); };
const addDays = (d, n) => { const r = new Date(d); r.setDate(r.getDate() + n); return r; };
const startOfWeek = (d) => { const r = new Date(d); r.setHours(0, 0, 0, 0); const day = (r.getDay() + 6) % 7; return addDays(r, -day); };
const startOfMonth = (d) => new Date(d.getFullYear(), d.getMonth(), 1);
const endOfMonth = (d) => new Date(d.getFullYear(), d.getMonth() + 1, 0);
const sameMonth = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth();
const fmtShort = (d) => d.toLocaleDateString("en-GB", { day: "2-digit", month: "short" });
const fmtDow = (d) => d.toLocaleDateString("en-GB", { weekday: "short" });
const fmtMonthYear = (d) => d.toLocaleDateString("en-GB", { month: "long", year: "numeric" });
const isWeekend = (d) => d.getDay() === 0 || d.getDay() === 6;
const isToday = (d) => toKey(d) === toKey(new Date());
/* ---------------------------------------------------------------------- */
/* Reporting period helpers — shared by Dashboard, Clients, Employees      */
/* ---------------------------------------------------------------------- */

const PERIOD_OPTIONS = [
  { key: "week", label: "Weekly" },
  { key: "month", label: "Monthly" },
  { key: "quarter", label: "Quarterly" },
  { key: "year", label: "Annual" },
];

function periodRange(period, anchor) {
  if (period === "week") { const s = startOfWeek(anchor); return [s, addDays(s, 6)]; }
  if (period === "quarter") {
    const q = Math.floor(anchor.getMonth() / 3);
    return [new Date(anchor.getFullYear(), q * 3, 1), new Date(anchor.getFullYear(), q * 3 + 3, 0)];
  }
  if (period === "year") return [new Date(anchor.getFullYear(), 0, 1), new Date(anchor.getFullYear(), 11, 31)];
  return [startOfMonth(anchor), endOfMonth(anchor)]; // "month" default
}
function shiftPeriodAnchor(period, anchor, dir) {
  if (period === "week") return addDays(anchor, dir * 7);
  if (period === "quarter") return new Date(anchor.getFullYear(), anchor.getMonth() + dir * 3, 1);
  if (period === "year") return new Date(anchor.getFullYear() + dir, anchor.getMonth(), 1);
  return new Date(anchor.getFullYear(), anchor.getMonth() + dir, 1); // "month"
}
function periodLabel(period, rangeStart, rangeEnd) {
  if (period === "week") return `${fmtShort(rangeStart)} – ${fmtShort(rangeEnd)}`;
  if (period === "quarter") return `Q${Math.floor(rangeStart.getMonth() / 3) + 1} ${rangeStart.getFullYear()}`;
  if (period === "year") return `${rangeStart.getFullYear()}`;
  return fmtMonthYear(rangeStart); // "month"
}
const LEAVE_TYPES = {
  annual: "Annual leave",
  sick: "Sick leave",
  medical: "Medical leave",
  other_legal: "Other (provided by law)",
  extra_acco: "Extra (from ACCO)",
};
// Types that require a note explaining the reason, since they cover
// statutory leave categories (marriage, bereavement, etc.).
const LEAVE_TYPES_REQUIRING_NOTE = new Set(["other_legal"]);
// Older data (from before these categories existed) used different keys —
// keep counting and labeling those correctly instead of showing "0"/blank.
const LEGACY_LEAVE_TYPE_MAP = { vacation: "annual", other: "other_legal" };
function normalizedLeaveType(type) { return LEGACY_LEAVE_TYPE_MAP[type] || type; }
function leaveTypeLabel(type) { return LEAVE_TYPES[normalizedLeaveType(type)] || type; }

function leaveStatusTone(status) {
  return status === "approved" ? "accent" : status === "rejected" ? "danger" : "warn";
}
/* ---------------------------------------------------------------------- */
/* Client revenue / productivity helpers                                   */
/* ---------------------------------------------------------------------- */

// Revenue a client generates for a given number of hours in a period.
// - Fixed fee (monthly retainer) is prorated for a weekly view.
// - Hourly rate applies to overage hours beyond the agreed allocation
//   when both a fixed fee and an allocation are set (retainer + overage
//   billing); otherwise it applies to all hours worked (pure hourly billing).
// Look up the historical record in effect on a given date: the latest
// entry whose effectiveDate is on or before dateKey. Falls back to the
// earliest entry if dateKey predates all history, and to `fallback` if
// there's no history at all.
function valueAsOf(history, dateKey, fallback) {
  if (!history || history.length === 0) return fallback;
  const sorted = [...history].sort((a, b) => a.effectiveDate.localeCompare(b.effectiveDate));
  let result = sorted[0];
  for (const h of sorted) {
    if (h.effectiveDate <= dateKey) result = h;
    else break;
  }
  return result;
}

// The earliest effective date recorded in a fee/cost history — used as
// "when did this client/employee's terms actually start applying",
// which the admin controls explicitly (e.g. backdating to when a
// long-standing client was first entered into the system), rather than
// whenever the database row happened to be created.
function earliestHistoryDate(history) {
  if (!history || history.length === 0) return null;
  return [...history].sort((a, b) => a.effectiveDate.localeCompare(b.effectiveDate))[0].effectiveDate;
}

// Sorts a history array chronologically and adds a derived "from"/"to"
// period to each entry — "to" is the day before the next entry starts,
// or "Present" for the most recent one. Purely for display.
function withPeriods(history) {
  const sorted = [...(history || [])].sort((a, b) => a.effectiveDate.localeCompare(b.effectiveDate));
  return sorted.map((h, i) => ({
    ...h,
    from: h.effectiveDate,
    to: i < sorted.length - 1 ? toKey(addDays(fromKey(sorted[i + 1].effectiveDate), -1)) : "Present",
  }));
}

// After editing or deleting a history entry, the employee/client's
// top-level "current" fields should track whichever entry is now the
// most recent by effective date.
function latestFromHistory(history, fieldNames) {
  if (!history || history.length === 0) return null;
  const sorted = [...history].sort((a, b) => a.effectiveDate.localeCompare(b.effectiveDate));
  const latest = sorted[sorted.length - 1];
  const result = {};
  fieldNames.forEach((f) => { result[f] = latest[f]; });
  return result;
}

function currentFeeOf(client) {
  return {
    fixedFee: client.fixedFee,
    allocation: client.allocation || { accounting: 0, tax: 0, payroll: 0, other: 0 },
    feeSplit: client.feeSplit || emptyFeeSplit(),
    feeSplitNote: client.feeSplitNote || "",
  };
}
// The fee terms (fixed fee / category allocation) in effect on a given
// date, e.g. the start of the period being reported on. Extra fees are
// tracked separately as one-off monthly entries — see extraFeeEntries.
function feeAsOf(client, dateKey) {
  return valueAsOf(client.feeHistory, dateKey, currentFeeOf(client));
}

/* ---------------------------------------------------------------------- */
/* Precise period math: splits a date range at every point the fee/cost    */
/* actually changed, and prorates each segment by its share of days —      */
/* so a mid-period change is reflected exactly, not just at period start.  */
/* ---------------------------------------------------------------------- */

const daysBetweenInclusive = (a, b) => Math.round((b - a) / 86400000) + 1;

// Splits [rangeStart, rangeEnd] into contiguous segments, one per distinct
// historical value in effect, using effectiveDate breakpoints that fall
// inside the range.
function regimesInRange(history, fallbackValue, rangeStart, rangeEnd) {
  const rangeStartKey = toKey(rangeStart);
  const rangeEndKey = toKey(rangeEnd);
  const sorted = history && history.length ? [...history].sort((a, b) => a.effectiveDate.localeCompare(b.effectiveDate)) : null;
  const breakpoints = sorted
    ? [...new Set(sorted.map((h) => h.effectiveDate).filter((d) => d > rangeStartKey && d <= rangeEndKey))].sort()
    : [];
  const boundaries = [rangeStartKey, ...breakpoints];
  return boundaries.map((segStartKey, i) => {
    const segEndKey = i + 1 < boundaries.length ? toKey(addDays(fromKey(boundaries[i + 1]), -1)) : rangeEndKey;
    return { start: fromKey(segStartKey), end: fromKey(segEndKey), value: valueAsOf(history, segStartKey, fallbackValue) };
  });
}

// Revenue AND allocation for a client over an arbitrary range, correctly
// prorated across any fee changes that happened mid-range.
// Revenue AND allocation for a client over an ARBITRARY date range (a
// week, month, quarter, year, whatever) — accrued day by day, using each
// day's own calendar-month length. This makes the math period-agnostic
// and exact even when a fee changed mid-range.
// Revenue, and category-by-category allocation vs. actual hours, for a
// client over an ARBITRARY date range (week, month, quarter, year, ...).
// Billing is a flat fixed fee + extra fees (no more hourly overage), each
// accrued day by day using each day's own calendar-month length so a
// mid-range fee change is reflected exactly and the math is period-agnostic.
// Sum of a client's one-off monthly extra fees that fall within a date
// range, accrued day by day (each entry's amount is spread evenly across
// the days of the month it was billed for) — so it works for any period
// length and any mix of months.
// Same accrual, but broken down by the cost center each extra fee was
// tagged with when it was added.
function extraFeesByCostCenterForRange(client, rangeStart, rangeEnd) {
  const endDate = client.endDate ? fromKey(client.endDate) : null;
  const byCC = emptyByCC();
  let d = new Date(rangeStart);
  while (d <= rangeEnd) {
    if (!endDate || d <= endDate) {
      const monthKey = `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
      const daysInThisMonth = endOfMonth(d).getDate();
      (client.extraFeeEntries || []).filter((e) => e.month === monthKey).forEach((e) => {
        const cc = COST_CENTERS.includes(e.costCenter) ? e.costCenter : "other";
        byCC[cc] += (Number(e.amount) || 0) / daysInThisMonth;
      });
    }
    d = addDays(d, 1);
  }
  return byCC;
}
function extraFeesAccruedForRange(client, rangeStart, rangeEnd) {
  const byCC = extraFeesByCostCenterForRange(client, rangeStart, rangeEnd);
  return Object.values(byCC).reduce((s, v) => s + v, 0);
}

function periodClientMetrics(client, periodEntries, rangeStart, rangeEnd) {
  const startDateKey = earliestHistoryDate(client.feeHistory);
  const startDate = startDateKey ? fromKey(startDateKey) : null;
  const endDate = client.endDate ? fromKey(client.endDate) : null;
  const regimes = regimesInRange(client.feeHistory, currentFeeOf(client), rangeStart, rangeEnd);
  // Revenue by cost center: extra fees carry their own tag; the fixed fee
  // follows the €-split recorded in the fee terms in effect on each day.
  // Any part of a fixed fee that has no split goes to "Unallocated" so the
  // totals still reconcile and the gap is visible on the dashboard.
  const revenueByCostCenter = extraFeesByCostCenterForRange(client, rangeStart, rangeEnd);
  let revenue = Object.values(revenueByCostCenter).reduce((s, v) => s + v, 0);
  const allocationByCategory = { accounting: 0, tax: 0, payroll: 0, other: 0 };
  let hasAllocation = false;
  regimes.forEach((seg) => {
    const fee = seg.value;
    const alloc = fee.allocation || {};
    const split = fee.feeSplit || {};
    const splitTotal = feeSplitTotal(split);
    let d = new Date(seg.start);
    while (d <= seg.end) {
      // Skip any day before the fee terms you set actually started
      // applying — controlled by the "Effective from" date you choose —
      // and any day after the client's end date, if one is set.
      if ((!startDate || d >= startDate) && (!endDate || d <= endDate)) {
        const daysInThisMonth = endOfMonth(d).getDate();
        if (fee.fixedFee) {
          revenue += fee.fixedFee / daysInThisMonth;
          COST_CENTERS.forEach((cc) => { if (split[cc]) revenueByCostCenter[cc] += split[cc] / daysInThisMonth; });
          const remainder = fee.fixedFee - splitTotal;
          if (remainder > 0.005) revenueByCostCenter[UNALLOCATED] += remainder / daysInThisMonth;
        }
        CATEGORIES.forEach((cat) => {
          if (alloc[cat]) { allocationByCategory[cat] += alloc[cat] / daysInThisMonth; hasAllocation = true; }
        });
      }
      d = addDays(d, 1);
    }
  });
  const actualByCategory = { accounting: 0, tax: 0, payroll: 0, other: 0 };
  periodEntries.forEach((e) => {
    const cat = CATEGORIES.includes(e.category) ? e.category : "other";
    actualByCategory[cat] += e.hours;
  });
  const allocation = hasAllocation ? CATEGORIES.reduce((s, c) => s + allocationByCategory[c], 0) : null;
  return { revenue, allocation, allocationByCategory, actualByCategory, revenueByCostCenter };
}

// Spreads an employee's labor cost for a period across the cost centers of
// their DEPARTMENT:
// - Accounting Team: by the categories of the hours they logged in the period
//   (accounting / tax / payroll / other);
// - otherwise (or with no hours logged): in proportion to the department's
//   revenue by cost center in the same period — e.g. Management Team cost is
//   split between Director and Finance the way their fees are;
// - if the department has neither hours nor revenue to go by: a single-cost-
//   center department gets it all, otherwise it stays "Unallocated" (still
//   counted in the department's own P&L).
function employeeCostByCostCenter(employee, cost, empPeriodEntries, deptRevenueByCC) {
  const byCC = emptyByCC();
  if (!(cost > 0)) return byCC;
  const dep = DEPARTMENTS.find((d) => d.key === (employee.department || DEFAULT_DEPARTMENT)) || DEPARTMENTS[0];
  if (dep.key === "accounting_team") {
    const chargeable = empPeriodEntries.filter((e) => !e.activity);
    const totalHours = chargeable.reduce((s, e) => s + e.hours, 0);
    if (totalHours > 0) {
      chargeable.forEach((e) => {
        const cat = CATEGORIES.includes(e.category) ? e.category : "other";
        byCC[cat] += cost * (e.hours / totalHours);
      });
      return byCC;
    }
  }
  const depRevenue = dep.costCenters.reduce((s, cc) => s + (deptRevenueByCC?.[cc] || 0), 0);
  if (depRevenue > 0) {
    dep.costCenters.forEach((cc) => { byCC[cc] += cost * ((deptRevenueByCC[cc] || 0) / depRevenue); });
  } else if (dep.costCenters.length === 1) {
    byCC[dep.costCenters[0]] += cost;
  } else {
    byCC[UNALLOCATED] += cost;
  }
  return byCC;
}

// Rental expenses that fall within a date range, accrued day by day (a
// recurring entry counts its monthly amount in every month from–to, an
// extra entry only in its own month; each month's amount is spread evenly
// over that month's days) —
// the same period-agnostic rule used for extra fees, so weekly, monthly,
// quarterly and annual views all reconcile.
function rentalExpensesForRange(expenses, rangeStart, rangeEnd) {
  const byCategory = Object.fromEntries(RENTAL_EXPENSE_CATEGORIES.map((k) => [k, 0]));
  const byProperty = {};
  let total = 0;
  let d = new Date(rangeStart);
  while (d <= rangeEnd) {
    const monthKey = monthKeyOf(d);
    const daysInThisMonth = endOfMonth(d).getDate();
    (expenses || []).filter((x) => rentalExpenseAppliesTo(x, monthKey)).forEach((x) => {
      const amt = (Number(x.amount) || 0) / daysInThisMonth;
      const cat = RENTAL_EXPENSE_CATEGORIES.includes(x.category) ? x.category : "rent";
      byCategory[cat] += amt;
      if (!byProperty[x.propertyId]) byProperty[x.propertyId] = Object.fromEntries([...RENTAL_EXPENSE_CATEGORIES, "total"].map((k) => [k, 0]));
      byProperty[x.propertyId][cat] += amt;
      byProperty[x.propertyId].total += amt;
      total += amt;
    });
    d = addDays(d, 1);
  }
  return { total, byCategory, byProperty };
}

// Labor cost for an employee over an arbitrary date range, accrued day by
// day the same way — correct across any salary/benefit change mid-range,
// and for any period length (week, month, quarter, year).
function periodEmployeeCost(employee, rangeStart, rangeEnd) {
  const regimes = regimesInRange(employee.costHistory, currentCostOf(employee), rangeStart, rangeEnd);
  let cost = 0;
  regimes.forEach((seg) => {
    const c = seg.value;
    const monthlyFull = (c.grossSalary || 0) + (c.socialSecurity || 0) + (c.ticketRestaurant || 0) + (c.insurance || 0) + (c.otherCost || 0);
    let d = new Date(seg.start);
    while (d <= seg.end) {
      cost += monthlyFull / endOfMonth(d).getDate();
      d = addDays(d, 1);
    }
  });
  return cost;
}
const fmtEur = (n) => "€" + n.toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 0 });
/* ---------------------------------------------------------------------- */
/* Greek public holidays & annual leave balance helpers                    */
/* ---------------------------------------------------------------------- */

const ANNUAL_LEAVE_DAYS = 25;
const LEAVE_TRACKING_START_YEAR = 2026; // balances carry over starting from this year

// Orthodox Easter (Gregorian date) via the Meeus Julian algorithm,
// converted from the Julian to the Gregorian calendar (+13 days,
// accurate for 1900–2099, which comfortably covers this app's use).
function orthodoxEaster(year) {
  const a = year % 4;
  const b = year % 7;
  const c = year % 19;
  const d = (19 * c + 15) % 30;
  const e = (2 * a + 4 * b - d + 34) % 7;
  const month = Math.floor((d + e + 114) / 31); // 3 = March, 4 = April (Julian)
  const day = ((d + e + 114) % 31) + 1;
  const julian = new Date(year, month - 1, day);
  return addDays(julian, 13);
}

// Fixed-date and Easter-linked Greek public holidays for a given year.
function getGreekPublicHolidays(year) {
  const easter = orthodoxEaster(year);
  const list = [
    { date: new Date(year, 0, 1), name: "New Year's Day" },
    { date: new Date(year, 0, 6), name: "Epiphany" },
    { date: addDays(easter, -48), name: "Clean Monday" },
    { date: new Date(year, 2, 25), name: "Independence Day" },
    { date: addDays(easter, -2), name: "Good Friday" },
    { date: addDays(easter, 1), name: "Easter Monday" },
    { date: new Date(year, 4, 1), name: "Labour Day" },
    { date: addDays(easter, 50), name: "Whit Monday" },
    { date: new Date(year, 7, 15), name: "Assumption Day" },
    { date: new Date(year, 9, 28), name: "Ohi Day" },
    { date: new Date(year, 11, 25), name: "Christmas Day" },
    { date: new Date(year, 11, 26), name: "Glorifying Mother of God" },
  ];
  return list;
}

// Holiday lookup map (dateKey -> name) spanning the given years.
function holidayMapForYears(years) {
  const map = new Map();
  years.forEach((y) => getGreekPublicHolidays(y).forEach((h) => map.set(toKey(h.date), h.name)));
  return map;
}

// Working days between two dates (inclusive), excluding weekends and
// the given set of public-holiday date keys.
function countWorkingDays(startKey, endKey, holidayMap) {
  let d = fromKey(startKey);
  const end = fromKey(endKey);
  let count = 0;
  while (d <= end) {
    if (!isWeekend(d) && !holidayMap.has(toKey(d))) count++;
    d = addDays(d, 1);
  }
  return count;
}

// Annual-leave days already used (approved, working days only) by an
// employee within a calendar year. Only the "annual" leave type counts
// against the entitlement — medical/other/extra do not.
function vacationDaysUsed(employeeId, year, leaveRequests, holidayMap) {
  return leaveRequests
    .filter((r) => r.employeeId === employeeId && normalizedLeaveType(r.type) === "annual" && r.status === "approved" && fromKey(r.startDate).getFullYear() === year)
    .reduce((sum, r) => sum + countWorkingDays(r.startDate, r.endDate, holidayMap), 0);
}

// Approved days used within a year, broken down by leave type — e.g.
// { annual: 12, sick: 2, medical: 0, other_legal: 1, extra_acco: 0 }.
function daysUsedByType(employeeId, year, leaveRequests, holidayMap) {
  const totals = {};
  Object.keys(LEAVE_TYPES).forEach((t) => { totals[t] = 0; });
  leaveRequests
    .filter((r) => r.employeeId === employeeId && r.status === "approved" && fromKey(r.startDate).getFullYear() === year)
    .forEach((r) => {
      const t = normalizedLeaveType(r.type);
      totals[t] = (totals[t] || 0) + countWorkingDays(r.startDate, r.endDate, holidayMap);
    });
  return totals;
}

// Cumulative annual-leave balance through the end of `uptoYear` — unused
// days (or a deficit) carry forward from every earlier tracked year.
// Positive = days still owed to the employee; negative = taken in advance.
function leaveBalanceThroughYear(employeeId, entitlement, uptoYear, leaveRequests) {
  let balance = 0;
  for (let y = LEAVE_TRACKING_START_YEAR; y <= uptoYear; y++) {
    const holidayMap = holidayMapForYears([y - 1, y, y + 1]);
    const used = vacationDaysUsed(employeeId, y, leaveRequests, holidayMap);
    balance += entitlement - used;
  }
  return balance;
}
/* ---------------------------------------------------------------------- */
/* Labor cost & productivity helpers                                       */
/* ---------------------------------------------------------------------- */

const MONTHLY_WORK_DAYS = 22; // standard basis for converting a monthly cost into an hourly rate

// Total monthly employer cost: gross salary + benefits + other.
function monthlyLaborCost(e) {
  return (e.grossSalary || 0) + (e.socialSecurity || 0) + (e.ticketRestaurant || 0) + (e.insurance || 0) + (e.otherCost || 0);
}

// Hourly cost, using a standard 22-working-day month and the employee's
// contracted weekly hours (converted to a daily figure over a 5-day week).
function hourlyLaborCost(e) {
  const dailyHours = (e.weeklyHours || 40) / 5;
  const monthlyHoursBasis = MONTHLY_WORK_DAYS * dailyHours;
  return monthlyHoursBasis > 0 ? monthlyLaborCost(e) / monthlyHoursBasis : 0;
}

function currentCostOf(e) {
  return { grossSalary: e.grossSalary || 0, socialSecurity: e.socialSecurity || 0, ticketRestaurant: e.ticketRestaurant || 0, insurance: e.insurance || 0, otherCost: e.otherCost || 0 };
}
// The cost breakdown in effect on a given date (e.g. the start of the
// period being reported on), for period-accurate Dashboard figures.
function costAsOf(e, dateKey) {
  return valueAsOf(e.costHistory, dateKey, currentCostOf(e));
}
function monthlyLaborCostAsOf(e, dateKey) {
  const c = costAsOf(e, dateKey);
  return (c.grossSalary || 0) + (c.socialSecurity || 0) + (c.ticketRestaurant || 0) + (c.insurance || 0) + (c.otherCost || 0);
}
function hourlyLaborCostAsOf(e, dateKey) {
  const dailyHours = (e.weeklyHours || 40) / 5;
  const basis = MONTHLY_WORK_DAYS * dailyHours;
  return basis > 0 ? monthlyLaborCostAsOf(e, dateKey) / basis : 0;
}

/* ---------------------------------------------------------------------- */
/* Small UI atoms                                                          */
/* ---------------------------------------------------------------------- */

function Btn({ children, onClick, variant = "primary", size = "md", icon: Icon, disabled, type = "button" }) {
  const base = {
    fontFamily: sans,
    fontWeight: 600,
    borderRadius: 6,
    cursor: disabled ? "not-allowed" : "pointer",
    display: "inline-flex",
    alignItems: "center",
    gap: 6,
    border: "1px solid transparent",
    transition: "background 120ms ease, border-color 120ms ease",
    opacity: disabled ? 0.5 : 1,
  };
  const sizes = { sm: { padding: "5px 10px", fontSize: 12.5 }, md: { padding: "8px 14px", fontSize: 13.5 } };
  const variants = {
    primary: { background: C.accent, color: "#fff" },
    secondary: { background: C.surface, color: C.ink, border: `1px solid ${C.border}` },
    ghost: { background: "transparent", color: C.inkMuted },
    danger: { background: C.dangerSoft, color: C.danger },
  };
  return (
    <button
      type={type}
      disabled={disabled}
      onClick={onClick}
      style={{ ...base, ...sizes[size], ...variants[variant] }}
    >
      {Icon && <Icon size={14} />}
      {children}
    </button>
  );
}

function Field({ label, children }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 12.5, color: C.inkMuted, fontFamily: sans }}>
      {label}
      {children}
    </label>
  );
}

const inputStyle = {
  fontFamily: sans,
  fontSize: 13.5,
  padding: "7px 9px",
  borderRadius: 6,
  border: `1px solid ${C.border}`,
  color: C.ink,
  background: C.surface,
};

function Badge({ children, tone = "neutral" }) {
  const tones = {
    neutral: { bg: C.border, fg: C.inkMuted },
    accent: { bg: C.accentSoft, fg: C.accentDark },
    warn: { bg: C.warnSoft, fg: C.warn },
    danger: { bg: C.dangerSoft, fg: C.danger },
  };
  const t = tones[tone];
  return (
    <span style={{
      fontFamily: sans, fontSize: 11, fontWeight: 600, padding: "2px 7px",
      borderRadius: 999, background: t.bg, color: t.fg, whiteSpace: "nowrap",
    }}>
      {children}
    </span>
  );
}

function Panel({ children, style }) {
  return (
    <div style={{
      background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10,
      padding: 18, ...style,
    }}>
      {children}
    </div>
  );
}

function StatCard({ label, value, sub, accent }) {
  return (
    <Panel style={{ flex: 1, minWidth: 150 }}>
      <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 8 }}>{label}</div>
      <div style={{ fontFamily: mono, fontSize: 26, fontWeight: 600, color: accent || C.ink }}>{value}</div>
      {sub && <div style={{ fontFamily: sans, fontSize: 12, color: C.inkFaint, marginTop: 4 }}>{sub}</div>}
    </Panel>
  );
}

function Modal({ title, onClose, children, width = 420 }) {
  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(18,33,30,0.35)", zIndex: 50,
      display: "flex", alignItems: "center", justifyContent: "center", padding: 16,
    }} onClick={onClose}>
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: C.surface, borderRadius: 12, width, maxWidth: "100%",
          padding: 22, border: `1px solid ${C.border}`, boxShadow: "0 12px 32px rgba(18,33,30,0.18)",
        }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h3 style={{ margin: 0, fontFamily: sans, fontSize: 16, fontWeight: 700, color: C.ink }}>{title}</h3>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: C.inkMuted, padding: 4 }}>
            <X size={18} />
          </button>
        </div>
        {children}
      </div>
    </div>
  );
}

function EmptyState({ icon: Icon, title, sub }) {
  return (
    <div style={{ textAlign: "center", padding: "40px 20px", color: C.inkMuted, fontFamily: sans }}>
      <Icon size={26} style={{ marginBottom: 10, color: C.inkFaint }} />
      <div style={{ fontWeight: 600, color: C.ink, marginBottom: 4 }}>{title}</div>
      {sub && <div style={{ fontSize: 13 }}>{sub}</div>}
    </div>
  );
}


function CenteredScreen({ children }) {
  return (
    <div style={{ minHeight: "100vh", background: C.sidebarBg, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: sans, padding: 20 }}>
      <FontImport />
      {children}
    </div>
  );
}

function SetupScreen() {
  return (
    <CenteredScreen>
      <div style={{ width: 460, maxWidth: "100%", background: C.surface, borderRadius: 12, padding: 26, border: `1px solid ${C.border}` }}>
        <h2 style={{ marginTop: 0, fontFamily: sans, color: C.ink }}>Almost there</h2>
        <p style={{ fontFamily: sans, fontSize: 13.5, color: C.inkMuted, lineHeight: 1.6 }}>
          This app needs a Supabase project to store its data. Copy <code>.env.example</code> to{" "}
          <code>.env</code>, fill in your project URL and anon key, then restart the dev server
          (or redeploy). See <code>README.md</code> for the full setup steps.
        </p>
      </div>
    </CenteredScreen>
  );
}

function LoadingScreen({ text = "Loading…" }) {
  return <CenteredScreen><span style={{ color: "#fff", fontFamily: sans }}>{text}</span></CenteredScreen>;
}

function AuthScreen() {
  const [mode, setMode] = useState("signin"); // "signin" | "signup" | "forgot"
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [busy, setBusy] = useState(false);

  async function submit() {
    if (mode === "forgot") {
      if (!email.trim()) { setError("Enter your email."); return; }
      setError(""); setInfo(""); setBusy(true);
      const { error } = await supabase.auth.resetPasswordForEmail(email.trim(), { redirectTo: window.location.origin });
      setBusy(false);
      if (error) setError(error.message);
      else setInfo("If that email has an account, a reset link is on its way — check your inbox (and spam folder). The link may take a minute to arrive.");
      return;
    }
    if (!email.trim() || !password) { setError("Enter both your email and password."); return; }
    setError(""); setInfo(""); setBusy(true);
    if (mode === "signup") {
      const { data, error } = await supabase.auth.signUp({ email, password, options: { data: { name } } });
      if (error) setError(error.message);
      else if (!data.session) setInfo("Account created. Check your email to confirm it, then sign in.");
    } else {
      const { error } = await supabase.auth.signInWithPassword({ email, password });
      if (error) setError(error.message);
    }
    setBusy(false);
  }
  function handleKeyDown(e) {
    if (e.key === "Enter") submit();
  }
  function switchMode(next) {
    setMode(next); setError(""); setInfo("");
  }

  return (
    <CenteredScreen>
      <div style={{ width: 380, maxWidth: "100%" }}>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 8, marginBottom: 24 }}>
          <img src={LOGO_DATA_URI} alt="ACCO logo" style={{ width: 56, height: 56, objectFit: "contain" }} />
          <span style={{ color: "#fff", fontWeight: 700, fontSize: 20, letterSpacing: 0.5 }}>GLORIA</span>
          <span style={{ color: C.sidebarText, fontWeight: 500, fontSize: 12 }}>ACCOUNTING ACCO L.P.</span>
        </div>
        <div style={{ background: C.surface, borderRadius: 12, padding: 22, border: `1px solid ${C.border}`, display: "flex", flexDirection: "column", gap: 12 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 15, color: C.ink }}>
            {mode === "signin" ? "Sign in" : mode === "signup" ? "Create your account" : "Reset your password"}
          </div>
          {mode === "signup" && (
            <Field label="Full name">
              <input style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Maria Ioannou" />
            </Field>
          )}
          <Field label="Work email">
            <input type="text" style={inputStyle} value={email} onChange={(e) => setEmail(e.target.value)} onKeyDown={handleKeyDown} placeholder="name@acco.gr" />
          </Field>
          {mode !== "forgot" && (
            <Field label="Password">
              <input type="password" style={inputStyle} value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={handleKeyDown} placeholder="At least 6 characters" />
            </Field>
          )}
          {error && <div style={{ fontFamily: sans, fontSize: 12.5, color: C.danger, background: C.dangerSoft, padding: "8px 10px", borderRadius: 6 }}>{error}</div>}
          {info && <div style={{ fontFamily: sans, fontSize: 12.5, color: C.accentDark, background: C.accentSoft, padding: "8px 10px", borderRadius: 6 }}>{info}</div>}
          <Btn onClick={submit} disabled={busy}>
            {mode === "signin" ? "Sign in" : mode === "signup" ? "Create account" : "Send reset link"}
          </Btn>
          {mode === "signin" && (
            <button type="button" onClick={() => switchMode("forgot")}
              style={{ background: "none", border: "none", cursor: "pointer", color: C.inkMuted, fontFamily: sans, fontSize: 12.5, textAlign: "center" }}>
              Forgot password?
            </button>
          )}
          {mode !== "forgot" ? (
            <button type="button" onClick={() => switchMode(mode === "signin" ? "signup" : "signin")}
              style={{ background: "none", border: "none", cursor: "pointer", color: C.inkMuted, fontFamily: sans, fontSize: 12.5, textAlign: "center" }}>
              {mode === "signin" ? "Need an account? Sign up" : "Already have an account? Sign in"}
            </button>
          ) : (
            <button type="button" onClick={() => switchMode("signin")}
              style={{ background: "none", border: "none", cursor: "pointer", color: C.inkMuted, fontFamily: sans, fontSize: 12.5, textAlign: "center" }}>
              Back to sign in
            </button>
          )}
        </div>
      </div>
    </CenteredScreen>
  );
}

function UpdatePasswordScreen({ onDone }) {
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState(false);

  async function submit() {
    if (password.length < 6) { setError("Password must be at least 6 characters."); return; }
    if (password !== confirm) { setError("Passwords don't match."); return; }
    setError(""); setBusy(true);
    const { error } = await supabase.auth.updateUser({ password });
    setBusy(false);
    if (error) { setError(error.message); return; }
    setDone(true);
  }
  function handleKeyDown(e) {
    if (e.key === "Enter") submit();
  }

  return (
    <CenteredScreen>
      <div style={{ width: 380, maxWidth: "100%" }}>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 8, marginBottom: 24 }}>
          <img src={LOGO_DATA_URI} alt="ACCO logo" style={{ width: 56, height: 56, objectFit: "contain" }} />
          <span style={{ color: "#fff", fontWeight: 700, fontSize: 20, letterSpacing: 0.5 }}>GLORIA</span>
          <span style={{ color: C.sidebarText, fontWeight: 500, fontSize: 12 }}>ACCOUNTING ACCO L.P.</span>
        </div>
        <div style={{ background: C.surface, borderRadius: 12, padding: 22, border: `1px solid ${C.border}`, display: "flex", flexDirection: "column", gap: 12 }}>
          {done ? (
            <>
              <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 15, color: C.ink }}>Password updated</div>
              <div style={{ fontFamily: sans, fontSize: 13, color: C.inkMuted }}>You can continue into HoursLedger now.</div>
              <Btn onClick={onDone}>Continue</Btn>
            </>
          ) : (
            <>
              <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 15, color: C.ink }}>Set a new password</div>
              <Field label="New password">
                <input type="password" autoFocus style={inputStyle} value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={handleKeyDown} placeholder="At least 6 characters" />
              </Field>
              <Field label="Confirm new password">
                <input type="password" style={inputStyle} value={confirm} onChange={(e) => setConfirm(e.target.value)} onKeyDown={handleKeyDown} placeholder="Repeat the password" />
              </Field>
              {error && <div style={{ fontFamily: sans, fontSize: 12.5, color: C.danger, background: C.dangerSoft, padding: "8px 10px", borderRadius: 6 }}>{error}</div>}
              <Btn onClick={submit} disabled={busy}>Update password</Btn>
            </>
          )}
        </div>
      </div>
    </CenteredScreen>
  );
}

/* ---------------------------------------------------------------------- */
/* Week strip control (shared)                                             */
/* ---------------------------------------------------------------------- */

function WeekNav({ weekStart, setWeekStart }) {
  const weekEnd = addDays(weekStart, 6);
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
      <button onClick={() => setWeekStart(addDays(weekStart, -7))} style={navBtnStyle}>
        <ChevronLeft size={16} />
      </button>
      <div style={{ fontFamily: sans, fontSize: 13.5, fontWeight: 600, color: C.ink, minWidth: 175, textAlign: "center" }}>
        {fmtShort(weekStart)} – {fmtShort(weekEnd)}, {weekEnd.getFullYear()}
      </div>
      <button onClick={() => setWeekStart(addDays(weekStart, 7))} style={navBtnStyle}>
        <ChevronRight size={16} />
      </button>
      <Btn variant="ghost" size="sm" onClick={() => setWeekStart(startOfWeek(new Date()))}>This week</Btn>
    </div>
  );
}
const navBtnStyle = {
  border: `1px solid ${C.border}`, background: C.surface, borderRadius: 6,
  width: 28, height: 28, display: "flex", alignItems: "center", justifyContent: "center",
  cursor: "pointer", color: C.ink,
};

// Shared week/month/quarter/year selector — used on Clients and Employees
// so their figures reflect a chosen reporting period rather than always
// "the current month".
function PeriodSelector({ period, setPeriod, anchor, setAnchor }) {
  const [rangeStart, rangeEnd] = periodRange(period, anchor);
  return (
    <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
      <div style={{ display: "flex", border: `1px solid ${C.border}`, borderRadius: 7, overflow: "hidden" }}>
        {PERIOD_OPTIONS.map((p) => (
          <button key={p.key} onClick={() => setPeriod(p.key)}
            style={{
              padding: "6px 10px", fontSize: 12, fontFamily: sans, fontWeight: 600, border: "none", cursor: "pointer",
              background: period === p.key ? C.accent : C.surface, color: period === p.key ? "#fff" : C.inkMuted,
            }}>
            {p.label}
          </button>
        ))}
      </div>
      <button onClick={() => setAnchor((a) => shiftPeriodAnchor(period, a, -1))} style={navBtnStyle}><ChevronLeft size={16} /></button>
      <div style={{ fontFamily: sans, fontSize: 13, fontWeight: 600, minWidth: 140, textAlign: "center" }}>
        {periodLabel(period, rangeStart, rangeEnd)}
      </div>
      <button onClick={() => setAnchor((a) => shiftPeriodAnchor(period, a, 1))} style={navBtnStyle}><ChevronRight size={16} /></button>
    </div>
  );
}

const iconBtnStyle = { background: "none", border: "none", cursor: "pointer", padding: 3, display: "flex" };

function PageHeader({ title, sub, right }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20, flexWrap: "wrap", gap: 10 }}>
      <div>
        <h1 style={{ fontFamily: sans, fontSize: 21, fontWeight: 700, color: C.ink, margin: 0 }}>{title}</h1>
        {sub && <div style={{ fontFamily: sans, fontSize: 13, color: C.inkMuted, marginTop: 4 }}>{sub}</div>}
      </div>
      {right}
    </div>
  );
}


function TableShell({ headers, children }) {
  return (
    <div className="scrollbar-thin" style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: sans, fontSize: 13 }}>
        <thead>
          <tr>
            {headers.map((h) => (
              <th key={h} style={{
                textAlign: "left", padding: "6px 8px", fontSize: 11.5, fontWeight: 600,
                color: C.inkFaint, borderBottom: `1px solid ${C.border}`,
              }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>{children}</tbody>
      </table>
    </div>
  );
}
function Td({ children, mono: isMono, right, style, title }) {
  return (
    <td title={title} style={{
      padding: "8px 8px", borderBottom: `1px solid ${C.border}`, color: C.ink,
      fontFamily: isMono ? mono : sans, textAlign: right ? "right" : "left", fontWeight: isMono ? 600 : 400,
      ...style,
    }}>{children}</td>
  );
}


function Sidebar({ user, view, setView }) {
  const employeeNav = [
    { key: "timesheet", label: "My Timesheet", icon: Clock },
    { key: "timeoff", label: "Time Off", icon: CalendarDays },
    { key: "calendar", label: "Calendar", icon: CalendarRange },
  ];
  const adminNav = [
    { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { key: "timesheet", label: "My Timesheet", icon: Clock },
    { key: "timesheets", label: "All Timesheets", icon: Clock },
    { key: "timeoff", label: "Time Off", icon: CalendarDays },
    { key: "leaveapprovals", label: "Leave Approvals", icon: ClipboardCheck },
    { key: "calendar", label: "Calendar", icon: CalendarRange },
    { key: "clients", label: "Clients", icon: Building2 },
    { key: "employees", label: "Employees", icon: Users },
    { key: "rental", label: "Rental", icon: Home },
    { key: "reports", label: "Reports", icon: SlidersHorizontal },
  ];
  const items = user.role === "admin" ? adminNav : employeeNav;
  return (
    <div style={{ width: 224, minWidth: 224, background: C.sidebarBg, minHeight: "100vh", display: "flex", flexDirection: "column", padding: "20px 14px", fontFamily: sans }}>
      <div style={{ display: "flex", alignItems: "center", gap: 9, padding: "0 6px", marginBottom: 26 }}>
        <img src={LOGO_DATA_URI} alt="ACCO logo" style={{ width: 28, height: 28, objectFit: "contain" }} />
        <div style={{ display: "flex", flexDirection: "column", lineHeight: 1.15 }}>
          <span style={{ color: "#fff", fontWeight: 700, fontSize: 15.5 }}>GLORIA</span>
          <span style={{ color: C.sidebarText, fontWeight: 500, fontSize: 9.5 }}>ACCOUNTING ACCO L.P.</span>
        </div>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 3, flex: 1 }}>
        {items.map((it) => {
          const active = view === it.key;
          return (
            <button key={it.key} onClick={() => setView(it.key)}
              style={{
                display: "flex", alignItems: "center", gap: 10, padding: "9px 11px", borderRadius: 7,
                border: "none", cursor: "pointer", textAlign: "left",
                background: active ? C.sidebarBgActive : "transparent",
                color: active ? C.sidebarTextActive : C.sidebarText,
                fontSize: 13.5, fontWeight: active ? 600 : 500,
              }}>
              <it.icon size={16} />
              {it.label}
            </button>
          );
        })}
      </div>
      <div style={{ borderTop: `1px solid rgba(255,255,255,0.08)`, paddingTop: 14, marginTop: 10 }}>
        <div style={{ color: "#fff", fontSize: 13, fontWeight: 600 }}>{user.name}</div>
        <div style={{ color: C.sidebarText, fontSize: 11.5, marginBottom: 10 }}>
          {user.title ? `${user.title} · ${user.role === "admin" ? "Manager" : "Employee"}` : (user.role === "admin" ? "Manager" : "Employee")}
        </div>
        <button onClick={() => supabase.auth.signOut()}
          style={{ display: "flex", alignItems: "center", gap: 7, background: "transparent", border: `1px solid rgba(255,255,255,0.12)`, color: C.sidebarText, borderRadius: 7, padding: "7px 10px", fontSize: 12.5, cursor: "pointer", width: "100%" }}>
          <LogOut size={13} /> Sign out
        </button>
      </div>
    </div>
  );
}


/* ---------------------------------------------------------------------- */
/* Supabase table hook: fetches once, then stays live via realtime         */
/* ---------------------------------------------------------------------- */

// Supabase returns at most 1,000 rows per request. Tables like time_entries
// grow past that within weeks, and without paging the newest rows would
// silently drop out of the app — so every table is fetched in pages until
// a short page comes back. Rows are ordered by a stable key so pages don't
// overlap or skip.
const PAGE_SIZE = 1000;
const STABLE_ORDER = { locked_weeks: ["employee_id", "week_start"], app_settings: ["key"] };
function useTable(table, mapRow, orderColumn) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);

  const refetch = useCallback(async () => {
    if (!supabase) { setLoading(false); return; }
    const orderKeys = orderColumn ? [orderColumn, "id"] : (STABLE_ORDER[table] || ["id"]);
    const all = [];
    let from = 0;
    while (true) {
      let q = supabase.from(table).select("*").range(from, from + PAGE_SIZE - 1);
      orderKeys.forEach((k) => { q = q.order(k, { ascending: true }); });
      const { data, error } = await q;
      if (error) { console.error(`Couldn't load ${table}:`, error.message); setLoading(false); return; }
      all.push(...(data || []));
      if (!data || data.length < PAGE_SIZE) break;
      from += PAGE_SIZE;
    }
    setRows(all.map(mapRow));
    setLoading(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [table]);

  useEffect(() => {
    refetch();
    if (!supabase) return;
    const channel = supabase
      .channel(`${table}_sync_${Math.random().toString(36).slice(2)}`)
      .on("postgres_changes", { event: "*", schema: "public", table }, () => refetch())
      .subscribe();
    return () => { supabase.removeChannel(channel); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refetch]);

  return [rows, refetch, loading];
}

const mapProfile = (r) => ({ id: r.id, name: r.name, email: r.email, title: r.title, role: r.role, weeklyHours: Number(r.weekly_hours), annualLeaveDays: Number(r.annual_leave_days), active: r.active, createdAt: r.created_at ? r.created_at.slice(0, 10) : null, department: r.department || DEFAULT_DEPARTMENT, endDate: r.end_date || null });
// Same idea as clientBilledInRange: "active" only controls whether someone
// can be picked for new work; whether their cost counts in a period depends
// on when their cost history starts and, if they've left, their departure date.
function employeeBilledInRange(e, rStart, rEnd) {
  const start = earliestHistoryDate(e.costHistory);
  if (start && fromKey(start) > rEnd) return false;
  if (e.endDate && fromKey(e.endDate) < rStart) return false;
  return true;
}
function clientBilledInRange(c, rStart, rEnd) {
  const start = earliestHistoryDate(c.feeHistory);
  if (start && fromKey(start) > rEnd) return false;
  if (c.endDate && fromKey(c.endDate) < rStart) return false;
  return true;
}
// Company-wide revenue/cost/hours for one arbitrary date range — the
// minimal aggregate behind every growth/comparison figure (period-over-
// period, YoY). Deliberately simpler than the dashboard's per-client and
// per-employee breakdowns: it only needs the totals.
function computePeriodTotals(clients, employees, entries, rentalExpenses, rStart, rEnd) {
  const rangeEntries = entries.filter((e) => { const d = fromKey(e.date); return d >= rStart && d <= rEnd; });
  const revenue = clients.filter((c) => clientBilledInRange(c, rStart, rEnd))
    .reduce((sum, c) => sum + periodClientMetrics(c, rangeEntries.filter((e) => e.clientId === c.id), rStart, rEnd).revenue, 0);
  const laborCost = employees.filter((e) => employeeBilledInRange(e, rStart, rEnd))
    .reduce((sum, e) => sum + periodEmployeeCost(e, rStart, rEnd), 0);
  const cost = laborCost + rentalExpensesForRange(rentalExpenses, rStart, rEnd).total;
  const hours = rangeEntries.reduce((s, e) => s + e.hours, 0);
  const profit = revenue - cost;
  return { revenue, cost, profit, hours, margin: revenue > 0 ? (profit / revenue) * 100 : null };
}
// Percentage change from a base value — null when there's no base to
// compare against (so the UI can show "—" instead of a misleading number).
function pctChange(current, previous) {
  if (previous === null || previous === undefined || Math.abs(previous) < 0.005) return null;
  return ((current - previous) / Math.abs(previous)) * 100;
}
// The immediately preceding period of the same length, and the same
// calendar period one year earlier — used for period-over-period and YoY
// comparisons. Both are computed from the period's own start/end so they
// stay correct across week/month/quarter/year and custom ranges alike.
function comparisonRanges(rangeStart, rangeEnd) {
  const spanDays = Math.round((rangeEnd - rangeStart) / 86400000) + 1;
  const prevEnd = addDays(rangeStart, -1);
  const prevStart = addDays(prevEnd, -(spanDays - 1));
  const yoyStart = new Date(rangeStart.getFullYear() - 1, rangeStart.getMonth(), rangeStart.getDate());
  const yoyEnd = new Date(rangeEnd.getFullYear() - 1, rangeEnd.getMonth(), rangeEnd.getDate());
  return { prevStart, prevEnd, yoyStart, yoyEnd };
}
const mapCostHistory = (r) => ({ id: r.id, employeeId: r.employee_id, effectiveDate: r.effective_date, grossSalary: Number(r.gross_salary), socialSecurity: Number(r.social_security), ticketRestaurant: Number(r.ticket_restaurant), insurance: Number(r.insurance), otherCost: Number(r.other_cost) });
const mapClientRow = (r) => ({ id: r.id, name: r.name, fixedFee: r.fixed_fee === null ? null : Number(r.fixed_fee), endDate: r.end_date || null, active: r.active, createdAt: r.created_at ? r.created_at.slice(0, 10) : null });
const mapFeeHistory = (r) => ({
  id: r.id, clientId: r.client_id, effectiveDate: r.effective_date, fixedFee: r.fixed_fee === null ? null : Number(r.fixed_fee),
  allocation: { accounting: Number(r.alloc_accounting), tax: Number(r.alloc_tax), payroll: Number(r.alloc_payroll), other: Number(r.alloc_other) },
  feeSplit: Object.fromEntries(COST_CENTERS.map((cc) => [cc, Number(r[`cc_${cc}`]) || 0])),
  feeSplitNote: r.cc_other_note || "",
});
const mapExtraFee = (r) => ({ id: r.id, clientId: r.client_id, month: r.month, amount: Number(r.amount), note: r.note || "", costCenter: r.cost_center || "other" });
const mapEntry = (r) => ({ id: r.id, employeeId: r.employee_id, clientId: r.client_id || null, activity: r.activity || null, date: r.entry_date, hours: Number(r.hours), category: r.category, note: r.note || "" });
// Label for the "who/what" of an entry: the client's name, or the activity.
const entryTargetLabel = (e, clients) => e.activity ? (ACTIVITY_LABELS[e.activity] || e.activity) : (clients.find((c) => c.id === e.clientId)?.name || "—");
const mapLockedWeek = (r) => `${r.employee_id}|${r.week_start}`;
const mapRentalProperty = (r) => ({ id: r.id, name: r.name, address: r.address || "", note: r.note || "", active: r.active });
const mapRentalExpense = (r) => ({ id: r.id, propertyId: r.property_id, kind: r.kind === "recurring" ? "recurring" : "extra", month: r.month || null, fromMonth: r.from_month || null, toMonth: r.to_month || null, category: r.category, amount: Number(r.amount), note: r.note || "" });
const mapLeaveRequest = (r) => ({ id: r.id, employeeId: r.employee_id, startDate: r.start_date, endDate: r.end_date, type: r.type, note: r.note || "", status: r.status, decidedBy: r.decided_by, decidedAt: r.decided_at });

// A small shared key/value setting (e.g. Dashboard "Customize" choices),
// stored in app_settings and kept live the same way the tables above are.
function useAppSetting(key, defaultValue) {
  const [value, setValue] = useState(defaultValue);
  const [loaded, setLoaded] = useState(false);

  const refetch = useCallback(async () => {
    if (!supabase) { setLoaded(true); return; }
    const { data, error } = await supabase.from("app_settings").select("*").eq("key", key).maybeSingle();
    if (!error) setValue(data ? data.value : defaultValue);
    setLoaded(true);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [key]);

  useEffect(() => {
    refetch();
    if (!supabase) return;
    const channel = supabase
      .channel(`app_settings_${key}_${Math.random().toString(36).slice(2)}`)
      .on("postgres_changes", { event: "*", schema: "public", table: "app_settings", filter: `key=eq.${key}` }, () => refetch())
      .subscribe();
    return () => { supabase.removeChannel(channel); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refetch]);

  async function save(next) {
    setValue(next); // optimistic
    if (!supabase) return;
    await supabase.from("app_settings").upsert({ key, value: next, updated_at: new Date().toISOString() });
  }

  return [value, save, loaded];
}

// Builds an .xlsx workbook from one or more sheets and downloads it.
function exportToExcel(filename, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, rows }) => {
    const ws = XLSX.utils.json_to_sheet(rows && rows.length ? rows : [{}]);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  XLSX.writeFile(wb, filename);
}


/* ---------------------------------------------------------------------- */
/* Employee Timesheet view                                                 */
/* ---------------------------------------------------------------------- */

function EmployeeTimesheet({ user, clients, entries, refetchEntries, lockedWeeks, refetchLockedWeeks }) {
  const [weekStart, setWeekStart] = useState(startOfWeek(TODAY));
  const [form, setForm] = useState({ date: toKey(TODAY), clientId: clients.find((c) => c.active)?.id || "", category: "accounting", hours: "", note: "" });
  const [editingId, setEditingId] = useState(null);
  const [editHours, setEditHours] = useState("");
  const [editNote, setEditNote] = useState("");

  const weekKey = `${user.id}|${toKey(weekStart)}`;
  const isLocked = lockedWeeks.includes(weekKey);
  const days = Array.from({ length: 7 }, (_, i) => addDays(weekStart, i));

  // Moving between weeks (including past ones) also moves the entry date
  // into that week — same weekday — so backdated hours land where you're
  // looking. Any date can be typed in the Date field directly as well.
  function goToWeek(newStart) {
    setWeekStart(newStart);
    setForm((f) => {
      const cur = f.date ? fromKey(f.date) : TODAY;
      const dow = (cur.getDay() + 6) % 7; // Monday = 0
      return { ...f, date: toKey(addDays(newStart, dow)) };
    });
  }
  const formDate = form.date ? fromKey(form.date) : null;
  const formDateInWeek = formDate && formDate >= weekStart && formDate <= addDays(weekStart, 6);
  const formWeekLocked = formDate ? lockedWeeks.includes(`${user.id}|${toKey(startOfWeek(formDate))}`) : false;

  const myEntries = entries.filter((e) => e.employeeId === user.id);
  const weekEntries = myEntries.filter((e) => {
    const d = fromKey(e.date);
    return d >= weekStart && d <= addDays(weekStart, 6);
  });
  const monthEntries = myEntries.filter((e) => sameMonth(fromKey(e.date), weekStart));

  const weekTotal = weekEntries.reduce((s, e) => s + e.hours, 0);
  const monthTotal = monthEntries.reduce((s, e) => s + e.hours, 0);
  const weeklyTarget = user.weeklyHours || 40;
  const weeklyDiff = weekTotal - weeklyTarget;
  const onTargetWeek = Math.abs(weeklyDiff) < 0.01;
  const weeklyDiffLabel = onTargetWeek ? "On track" : weeklyDiff > 0 ? `+${weeklyDiff.toFixed(2)}h overtime` : `${weeklyDiff.toFixed(2)}h undertime`;
  const weeklyDiffColor = onTargetWeek ? C.accent : weeklyDiff > 0 ? C.warn : C.danger;

  const isActivity = isActivityValue(form.clientId);
  const activityKey = isActivity ? form.clientId.slice(ACTIVITY_PREFIX.length) : null;
  const otherNoteRequired = isActivity ? activityKey === "other_non_chargeable" : form.category === "other";
  const missingOtherNote = otherNoteRequired && !form.note.trim();

  async function addEntry() {
    if (!form.clientId || !form.hours || Number(form.hours) <= 0 || missingOtherNote) return;
    const d = fromKey(form.date);
    const wk = `${user.id}|${toKey(startOfWeek(d))}`;
    if (lockedWeeks.includes(wk)) {
      window.alert(`The week of ${fmtShort(startOfWeek(d))} is submitted and locked. Ask your manager to reopen it before adding hours there.`);
      return;
    }
    const { error } = await supabase.from("time_entries").insert({
      employee_id: user.id,
      client_id: isActivity ? null : form.clientId,
      activity: isActivity ? activityKey : null,
      category: isActivity ? "other" : form.category,
      entry_date: form.date, hours: Number(form.hours), note: form.note.trim() || null,
    });
    if (error) { window.alert(`Couldn't add entry: ${error.message}`); return; }
    setForm((f) => ({ ...f, hours: "", note: "" }));
    // If the hours were for another week (e.g. backdated), show that week.
    const wkStart = startOfWeek(d);
    if (toKey(wkStart) !== toKey(weekStart)) setWeekStart(wkStart);
    refetchEntries();
  }
  async function removeEntry(id) {
    await supabase.from("time_entries").delete().eq("id", id);
    refetchEntries();
  }
  async function saveEdit(id) {
    const val = Number(editHours);
    if (val > 0) {
      await supabase.from("time_entries").update({ hours: val, note: editNote.trim() || null }).eq("id", id);
      refetchEntries();
    }
    setEditingId(null);
  }
  async function lockWeek() {
    await supabase.from("locked_weeks").insert({ employee_id: user.id, week_start: toKey(weekStart) });
    refetchLockedWeeks();
  }

  return (
    <div>
      <PageHeader title="My Timesheet" sub="Log the hours you worked for each client, by day — or for internal, non-chargeable activities." />

      <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
        <StatCard label="This week" value={weekTotal.toFixed(2) + "h"} accent={C.accent} />
        <StatCard label={`vs. ${weeklyTarget}h/week target`} value={weeklyDiffLabel} accent={weeklyDiffColor} />
        <StatCard label={fmtMonthYear(weekStart)} value={monthTotal.toFixed(2) + "h"} />
        <StatCard label="Entries this week" value={weekEntries.length} />
      </div>

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
          <WeekNav weekStart={weekStart} setWeekStart={goToWeek} />
          {!isLocked ? (
            <Btn variant="primary" size="sm" icon={Lock} onClick={lockWeek}>Submit week</Btn>
          ) : (
            <Badge tone="warn">Submitted &amp; locked</Badge>
          )}
        </div>
        {isLocked && (
          <div style={{
            display: "flex", alignItems: "center", gap: 8, background: C.warnSoft, color: C.warn,
            padding: "8px 12px", borderRadius: 7, fontFamily: sans, fontSize: 12.5, marginBottom: 14,
          }}>
            <Lock size={13} /> This week is submitted and locked. Ask your manager to reopen it if you need to make changes.
          </div>
        )}

        {!isLocked && (
          <div style={{
            display: "flex", gap: 10, alignItems: "flex-end", marginBottom: 18, flexWrap: "wrap",
            borderBottom: `1px solid ${C.border}`, paddingBottom: 18,
          }}>
            <Field label="Date (any date, past weeks included)">
              <input type="date" style={{ ...inputStyle, width: 150, ...(formWeekLocked ? { border: `1px solid ${C.danger}` } : {}) }} value={form.date}
                onChange={(e) => setForm((f) => ({ ...f, date: e.target.value }))} />
            </Field>
            <Field label="Client / activity">
              <select style={{ ...inputStyle, width: 210 }} value={form.clientId}
                onChange={(e) => setForm((f) => ({ ...f, clientId: e.target.value }))}>
                <optgroup label="Clients">
                  {clients.filter((c) => c.active).map((c) => (
                    <option key={c.id} value={c.id}>{c.name}</option>
                  ))}
                </optgroup>
                <optgroup label="Internal — non-chargeable">
                  {ACTIVITIES.map((a) => (
                    <option key={a} value={ACTIVITY_PREFIX + a}>{ACTIVITY_LABELS[a]}</option>
                  ))}
                </optgroup>
              </select>
            </Field>
            {!isActivity && (
              <Field label="Category">
                <select style={{ ...inputStyle, width: 140 }} value={form.category}
                  onChange={(e) => setForm((f) => ({ ...f, category: e.target.value }))}>
                  {CATEGORIES.map((cat) => <option key={cat} value={cat}>{CATEGORY_LABELS[cat]}</option>)}
                </select>
              </Field>
            )}
            <Field label="Hours">
              <input type="number" min="0.25" step="0.25" placeholder="e.g. 3.5" style={{ ...inputStyle, width: 90 }}
                value={form.hours} onChange={(e) => setForm((f) => ({ ...f, hours: e.target.value }))} />
            </Field>
            <Field label={otherNoteRequired ? (isActivity ? "Note (required for \"Other non-chargeable\")" : "Note (required for \"Other\")") : "Note (optional)"}>
              <input type="text" placeholder={isActivity ? "e.g. Seminar on new tax law" : "e.g. VAT return review"} style={{ ...inputStyle, width: 200, ...(missingOtherNote ? { border: `1px solid ${C.danger}` } : {}) }}
                value={form.note} onChange={(e) => setForm((f) => ({ ...f, note: e.target.value }))} />
            </Field>
            <Btn icon={Plus} onClick={addEntry} disabled={missingOtherNote || formWeekLocked}>Add entry</Btn>
            {formWeekLocked && (
              <div style={{ width: "100%", fontFamily: sans, fontSize: 12, color: C.danger }}>
                The week containing {form.date} is submitted and locked — ask your manager to reopen it.
              </div>
            )}
            {!formWeekLocked && formDate && !formDateInWeek && (
              <div style={{ width: "100%", fontFamily: sans, fontSize: 12, color: C.inkMuted }}>
                This entry will be saved to the week of {fmtShort(startOfWeek(formDate))} and the view will jump there.
              </div>
            )}
          </div>
        )}

        {days.map((day) => {
          const dayEntries = weekEntries.filter((e) => e.date === toKey(day));
          const dayTotal = dayEntries.reduce((s, e) => s + e.hours, 0);
          return (
            <div key={toKey(day)} style={{ display: "flex", gap: 16, padding: "10px 0", borderBottom: `1px solid ${C.border}` }}>
              <div style={{ width: 90, flexShrink: 0 }}>
                <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 600, color: isToday(day) ? C.accent : C.ink }}>
                  {fmtDow(day)} {fmtShort(day)}
                </div>
                {isWeekend(day) && <div style={{ fontFamily: sans, fontSize: 10.5, color: C.inkFaint }}>Weekend</div>}
                {!isLocked && (
                  <button onClick={() => setForm((f) => ({ ...f, date: toKey(day) }))} title={`Add hours for ${fmtShort(day)}`}
                    style={{ ...iconBtnStyle, padding: 0, marginTop: 4, fontFamily: sans, fontSize: 11, color: form.date === toKey(day) ? C.accent : C.inkFaint, fontWeight: form.date === toKey(day) ? 700 : 500, alignItems: "center", gap: 3 }}>
                    <Plus size={11} /> {form.date === toKey(day) ? "Selected" : "Add hours"}
                  </button>
                )}
              </div>
              <div style={{ flex: 1 }}>
                {dayEntries.length === 0 ? (
                  <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkFaint }}>No hours logged</div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    {dayEntries.map((e) => (
                      <div key={e.id}>
                        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                          <span style={{ fontFamily: sans, fontSize: 13, color: C.ink, minWidth: 170 }}>{entryTargetLabel(e, clients)}</span>
                          {e.activity ? <Badge tone="warn">Non-chargeable</Badge> : <Badge tone="neutral">{CATEGORY_LABELS[e.category] || "Other"}</Badge>}
                          {editingId === e.id ? (
                            <>
                              <input type="number" min="0.25" step="0.25" autoFocus value={editHours}
                                onChange={(ev) => setEditHours(ev.target.value)}
                                style={{ ...inputStyle, width: 74, padding: "3px 7px", fontFamily: mono }} />
                              <input type="text" placeholder="Note (optional)" value={editNote}
                                onChange={(ev) => setEditNote(ev.target.value)}
                                style={{ ...inputStyle, width: 160, padding: "3px 7px" }} />
                              <button onClick={() => saveEdit(e.id)} style={iconBtnStyle}><Check size={14} color={C.accent} /></button>
                              <button onClick={() => setEditingId(null)} style={iconBtnStyle}><X size={14} color={C.inkMuted} /></button>
                            </>
                          ) : (
                            <>
                              <span style={{ fontFamily: mono, fontSize: 13, color: C.ink, fontWeight: 600 }}>{e.hours.toFixed(2)}h</span>
                              {!isLocked && (
                                <>
                                  <button onClick={() => { setEditingId(e.id); setEditHours(String(e.hours)); setEditNote(e.note || ""); }} style={iconBtnStyle}>
                                    <Pencil size={13} color={C.inkMuted} />
                                  </button>
                                  <button onClick={() => removeEntry(e.id)} style={iconBtnStyle}>
                                    <Trash2 size={13} color={C.inkMuted} />
                                  </button>
                                </>
                              )}
                            </>
                          )}
                        </div>
                        {editingId !== e.id && e.note && (
                          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkFaint, marginTop: 2, marginLeft: 0 }}>{e.note}</div>
                        )}
                      </div>
                    ))}
                  </div>
                )}
              </div>
              <div style={{ width: 60, textAlign: "right", fontFamily: mono, fontSize: 13, fontWeight: 600, color: dayTotal > 0 ? C.ink : C.inkFaint }}>
                {dayTotal > 0 ? dayTotal.toFixed(2) + "h" : "—"}
              </div>
            </div>
          );
        })}
        <div style={{ display: "flex", justifyContent: "flex-end", paddingTop: 12, fontFamily: sans, fontSize: 13, fontWeight: 700, color: C.ink }}>
          Week total:&nbsp;<span style={{ fontFamily: mono, color: C.accent }}>{weekTotal.toFixed(2)}h</span>
        </div>
      </Panel>
    </div>
  );
}


/* ---------------------------------------------------------------------- */
/* Time Off — request (employee) and history                              */
/* ---------------------------------------------------------------------- */

function TimeOff({ user, employees, leaveRequests, refetchLeaveRequests }) {
  const [form, setForm] = useState({ startDate: toKey(TODAY), endDate: toKey(TODAY), type: "annual", note: "" });

  const mine = leaveRequests
    .filter((r) => r.employeeId === user.id)
    .sort((a, b) => b.startDate.localeCompare(a.startDate));

  const requestYear = form.startDate ? fromKey(form.startDate).getFullYear() : TODAY.getFullYear();
  const holidayMap = holidayMapForYears([requestYear - 1, requestYear, requestYear + 1]);
  const usedThisYear = vacationDaysUsed(user.id, requestYear, leaveRequests, holidayMap);
  const annualLeaveDays = user.annualLeaveDays || ANNUAL_LEAVE_DAYS;
  const remaining = leaveBalanceThroughYear(user.id, annualLeaveDays, requestYear, leaveRequests);

  const validRange = form.startDate && form.endDate && form.endDate >= form.startDate;
  const requestedDays = validRange ? countWorkingDays(form.startDate, form.endDate, holidayMap) : 0;
  const overBalance = form.type === "annual" && requestedDays > remaining;
  const noteRequired = LEAVE_TYPES_REQUIRING_NOTE.has(form.type);
  const missingRequiredNote = noteRequired && !form.note.trim();
  const usageByType = daysUsedByType(user.id, requestYear, leaveRequests, holidayMap);
  const carryoverIn = leaveBalanceThroughYear(user.id, annualLeaveDays, requestYear - 1, leaveRequests);

  async function submit() {
    if (!validRange || overBalance || missingRequiredNote) return;
    await supabase.from("leave_requests").insert({
      employee_id: user.id, start_date: form.startDate, end_date: form.endDate,
      type: form.type, note: form.note.trim() || null, status: "pending",
    });
    setForm((f) => ({ ...f, note: "" }));
    refetchLeaveRequests();
  }
  async function cancelRequest(id) {
    await supabase.from("leave_requests").delete().eq("id", id);
    refetchLeaveRequests();
  }
  const decidedByName = (id) => employees.find((e) => e.id === id)?.name || "—";

  return (
    <div>
      <PageHeader title="Time Off" sub="Request leave and track approval status." />
      <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
        <StatCard label={`Annual leave used in ${requestYear}`} value={`${usedThisYear} / ${annualLeaveDays}`} sub="working days" />
        <StatCard label={`Carried over from ${requestYear - 1}`} value={`${carryoverIn >= 0 ? "+" : ""}${carryoverIn}d`} accent={carryoverIn < 0 ? C.danger : undefined} />
        <StatCard label={`Balance through ${requestYear}`} value={`${remaining >= 0 ? "+" : ""}${remaining}d`} accent={remaining < 0 ? C.danger : remaining === 0 ? C.warn : C.accent} sub="carries over year to year" />
      </div>
      <div style={{ display: "flex", gap: 18, flexWrap: "wrap" }}>
        <Panel style={{ flex: "0 0 320px" }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 14 }}>New request</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Field label="Start date (past dates allowed — e.g. leave already taken)">
              <input type="date" style={inputStyle} value={form.startDate} onChange={(e) => setForm((f) => ({ ...f, startDate: e.target.value, endDate: f.endDate < e.target.value ? e.target.value : f.endDate }))} />
            </Field>
            <Field label="End date">
              <input type="date" style={inputStyle} value={form.endDate} onChange={(e) => setForm((f) => ({ ...f, endDate: e.target.value }))} />
            </Field>
            {!validRange && form.startDate && form.endDate && (
              <div style={{ fontFamily: sans, fontSize: 12, color: C.danger }}>End date must be the same as or after the start date.</div>
            )}
            <Field label="Type">
              <select style={inputStyle} value={form.type} onChange={(e) => setForm((f) => ({ ...f, type: e.target.value }))}>
                {Object.entries(LEAVE_TYPES).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
              </select>
            </Field>
            <Field label={noteRequired ? "Note (required for this leave type)" : "Note (optional)"}>
              <input style={{ ...inputStyle, ...(missingRequiredNote ? { border: `1px solid ${C.danger}` } : {}) }} value={form.note} onChange={(e) => setForm((f) => ({ ...f, note: e.target.value }))} placeholder="e.g. Family trip" />
            </Field>
            {noteRequired && (
              <div style={{ fontFamily: sans, fontSize: 12, color: missingRequiredNote ? C.danger : C.inkMuted }}>
                This leave type requires a note explaining the reason.
              </div>
            )}
            {validRange && form.type === "annual" && (
              <div style={{ fontFamily: sans, fontSize: 12, color: overBalance ? C.danger : C.inkMuted }}>
                Uses {requestedDays} working day{requestedDays === 1 ? "" : "s"} of your {remaining}d balance
                {overBalance ? " — exceeds your balance." : "."}
              </div>
            )}
            <Btn icon={Plus} onClick={submit} disabled={!validRange || overBalance || missingRequiredNote}>Submit request</Btn>
          </div>
        </Panel>
        <Panel style={{ flex: "0 0 220px" }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 4 }}>Used in {requestYear}</div>
          <div style={{ fontFamily: sans, fontSize: 11.5, color: C.inkFaint, marginBottom: 10 }}>By leave type, approved only</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {Object.entries(LEAVE_TYPES).map(([key, label]) => (
              <div key={key} style={{ display: "flex", justifyContent: "space-between", fontFamily: sans, fontSize: 12.5 }}>
                <span style={{ color: C.inkMuted }}>{label}</span>
                <span style={{ fontFamily: mono, fontWeight: 600, color: C.ink }}>{usageByType[key] || 0}d</span>
              </div>
            ))}
          </div>
        </Panel>
        <Panel style={{ flex: 1, minWidth: 320 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 12 }}>Your requests</div>
          {mine.length === 0 ? <EmptyState icon={CalendarDays} title="No time off requested yet" /> : (
            <TableShell headers={["Dates", "Type", "Note", "Status", ""]}>
              {mine.map((r) => (
                <tr key={r.id}>
                  <Td>{fmtShort(fromKey(r.startDate))} – {fmtShort(fromKey(r.endDate))}</Td>
                  <Td>{leaveTypeLabel(r.type)}</Td>
                  <Td>{r.note || "—"}</Td>
                  <Td>
                    <Badge tone={leaveStatusTone(r.status)}>
                      {r.status === "approved" ? `Approved by ${decidedByName(r.decidedBy)}` : r.status === "rejected" ? `Rejected by ${decidedByName(r.decidedBy)}` : "Pending"}
                    </Badge>
                  </Td>
                  <Td right>
                    {r.status === "pending" && (
                      <Btn size="sm" variant="ghost" icon={Trash2} onClick={() => cancelRequest(r.id)}>Cancel</Btn>
                    )}
                  </Td>
                </tr>
              ))}
            </TableShell>
          )}
        </Panel>
      </div>
    </div>
  );
}


/* ---------------------------------------------------------------------- */
/* Admin: Leave Approvals                                                  */
/* ---------------------------------------------------------------------- */

function LeaveApprovals({ user, employees, leaveRequests, refetchLeaveRequests }) {
  const [statusFilter, setStatusFilter] = useState("pending");
  const [year, setYear] = useState(TODAY.getFullYear());
  const empName = (id) => employees.find((e) => e.id === id)?.name || "—";

  const holidayMap = holidayMapForYears([year - 1, year, year + 1]);
  const balances = employees.filter((e) => e.active).map((e) => {
    const entitlement = e.annualLeaveDays || ANNUAL_LEAVE_DAYS;
    const used = vacationDaysUsed(e.id, year, leaveRequests, holidayMap);
    const carryover = leaveBalanceThroughYear(e.id, entitlement, year - 1, leaveRequests);
    const balance = carryover + entitlement - used;
    const byType = daysUsedByType(e.id, year, leaveRequests, holidayMap);
    return { ...e, entitlement, used, carryover, balance, byType };
  }).sort((a, b) => a.balance - b.balance);

  const rows = leaveRequests
    .filter((r) => statusFilter === "all" || r.status === statusFilter)
    .sort((a, b) => b.startDate.localeCompare(a.startDate));

  async function decide(id, status) {
    await supabase.from("leave_requests").update({ status, decided_by: user.id, decided_at: toKey(TODAY) }).eq("id", id);
    refetchLeaveRequests();
  }

  return (
    <div>
      <PageHeader title="Leave Approvals" sub="Review and decide on time off requests from your team." />

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 4 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink }}>Annual leave balances</div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <button onClick={() => setYear((y) => y - 1)} style={navBtnStyle}><ChevronLeft size={16} /></button>
            <div style={{ fontFamily: sans, fontSize: 13, fontWeight: 600, minWidth: 50, textAlign: "center" }}>{year}</div>
            <button onClick={() => setYear((y) => y + 1)} style={navBtnStyle}><ChevronRight size={16} /></button>
          </div>
        </div>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>
          Balance carries over from year to year — a negative number means more days were taken than earned. Public holidays and weekends don't count against it.
        </div>
        <TableShell headers={["Employee", `Carried over from ${year - 1}`, "Entitlement", `Used in ${year}`, `Balance through ${year}`, ""]}>
          {balances.map((e) => {
            const pct = Math.min(100, (e.used / e.entitlement) * 100);
            const tone = e.balance < 0 ? "danger" : e.balance <= 5 ? "warn" : "accent";
            return (
              <tr key={e.id}>
                <Td>{e.name}</Td>
                <Td mono style={{ color: e.carryover < 0 ? C.danger : undefined }}>{e.carryover >= 0 ? "+" : ""}{e.carryover}d</Td>
                <Td mono>{e.entitlement}d</Td>
                <Td mono>{e.used}d</Td>
                <Td mono style={{ color: e.balance < 0 ? C.danger : undefined, fontWeight: 700 }}>{e.balance >= 0 ? "+" : ""}{e.balance}d</Td>
                <Td>
                  <div style={{ width: 120, height: 6, borderRadius: 999, background: C.border, overflow: "hidden" }}>
                    <div style={{ width: `${pct}%`, height: "100%", background: tone === "danger" ? C.danger : tone === "warn" ? C.warn : C.accent }} />
                  </div>
                </Td>
              </tr>
            );
          })}
        </TableShell>
      </Panel>

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 4 }}>Leave usage by type — {year}</div>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Approved working days, per employee, per category.</div>
        <TableShell headers={["Employee", ...Object.values(LEAVE_TYPES)]}>
          {balances.map((e) => (
            <tr key={e.id}>
              <Td>{e.name}</Td>
              {Object.keys(LEAVE_TYPES).map((key) => (
                <Td key={key} mono>{e.byType[key] || 0}d</Td>
              ))}
            </tr>
          ))}
        </TableShell>
      </Panel>

      <Panel>
        <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          {["pending", "approved", "rejected", "all"].map((s) => (
            <button key={s} onClick={() => setStatusFilter(s)}
              style={{
                padding: "6px 12px", borderRadius: 999, border: `1px solid ${C.border}`, cursor: "pointer",
                fontFamily: sans, fontSize: 12.5, fontWeight: 600, textTransform: "capitalize",
                background: statusFilter === s ? C.accent : C.surface, color: statusFilter === s ? "#fff" : C.inkMuted,
              }}>
              {s}
            </button>
          ))}
        </div>
        {rows.length === 0 ? <EmptyState icon={ClipboardCheck} title="No requests here" /> : (
          <TableShell headers={["Employee", "Dates", "Type", "Note", "Status", ""]}>
            {rows.map((r) => (
              <tr key={r.id}>
                <Td>{empName(r.employeeId)}</Td>
                <Td>{fmtShort(fromKey(r.startDate))} – {fmtShort(fromKey(r.endDate))}</Td>
                <Td>{leaveTypeLabel(r.type)}</Td>
                <Td>{r.note || "—"}</Td>
                <Td><Badge tone={leaveStatusTone(r.status)}>{r.status}</Badge></Td>
                <Td right>
                  {r.status === "pending" ? (
                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      <Btn size="sm" variant="secondary" icon={Check} onClick={() => decide(r.id, "approved")}>Approve</Btn>
                      <Btn size="sm" variant="danger" icon={X} onClick={() => decide(r.id, "rejected")}>Reject</Btn>
                    </div>
                  ) : (
                    <span style={{ fontFamily: sans, fontSize: 12, color: C.inkFaint }}>by {empName(r.decidedBy)}</span>
                  )}
                </Td>
              </tr>
            ))}
          </TableShell>
        )}
      </Panel>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Company Calendar — public holidays + approved leave                    */
/* ---------------------------------------------------------------------- */

const WEEKDAY_LABELS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

function buildMonthGrid(monthAnchor) {
  const first = startOfMonth(monthAnchor);
  const gridStart = startOfWeek(first);
  const days = [];
  for (let i = 0; i < 42; i++) days.push(addDays(gridStart, i));
  return days;
}

function TeamCalendar({ employees, leaveRequests }) {
  const [monthAnchor, setMonthAnchor] = useState(startOfMonth(TODAY));
  const year = monthAnchor.getFullYear();
  const holidayMap = holidayMapForYears([year - 1, year, year + 1]);
  const days = buildMonthGrid(monthAnchor);
  const empName = (id) => employees.find((e) => e.id === id)?.name || "—";
  const initials = (id) => (empName(id).match(/\b\w/g) || []).slice(0, 2).join("").toUpperCase();

  const approvedLeave = leaveRequests.filter((r) => r.status === "approved");
  function leaveOn(day) {
    const k = toKey(day);
    return approvedLeave.filter((r) => r.startDate <= k && k <= r.endDate);
  }

  return (
    <div>
      <PageHeader title="Calendar" sub="Public holidays and approved time off, at a glance."
        right={
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <button onClick={() => setMonthAnchor(new Date(year, monthAnchor.getMonth() - 1, 1))} style={navBtnStyle}><ChevronLeft size={16} /></button>
            <div style={{ fontFamily: sans, fontSize: 13.5, fontWeight: 600, minWidth: 140, textAlign: "center" }}>{fmtMonthYear(monthAnchor)}</div>
            <button onClick={() => setMonthAnchor(new Date(year, monthAnchor.getMonth() + 1, 1))} style={navBtnStyle}><ChevronRight size={16} /></button>
            <Btn variant="ghost" size="sm" onClick={() => setMonthAnchor(startOfMonth(new Date()))}>This month</Btn>
          </div>
        } />
      <Panel>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 6, marginBottom: 6 }}>
          {WEEKDAY_LABELS.map((d) => (
            <div key={d} style={{ fontFamily: sans, fontSize: 11, fontWeight: 600, color: C.inkFaint, textAlign: "center", padding: "2px 0" }}>{d}</div>
          ))}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 6 }}>
          {days.map((day) => {
            const inMonth = day.getMonth() === monthAnchor.getMonth();
            const holiday = holidayMap.get(toKey(day));
            const dayLeave = leaveOn(day);
            return (
              <div key={toKey(day)} style={{
                minHeight: 78, borderRadius: 8, padding: 6,
                border: `1px solid ${C.border}`,
                background: holiday ? C.warnSoft : isToday(day) ? C.accentSoft : C.surface,
                opacity: inMonth ? 1 : 0.4,
              }}>
                <div style={{ fontFamily: mono, fontSize: 11.5, fontWeight: 600, color: isToday(day) ? C.accentDark : C.ink }}>{day.getDate()}</div>
                {holiday && <div style={{ fontFamily: sans, fontSize: 10, color: C.warn, marginTop: 2, lineHeight: 1.3 }}>{holiday}</div>}
                {dayLeave.length > 0 && (
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 3, marginTop: 4 }}>
                    {dayLeave.map((r) => (
                      <span key={r.id} title={`${empName(r.employeeId)} — ${leaveTypeLabel(r.type)}`} style={{
                        fontFamily: sans, fontSize: 9.5, fontWeight: 700, color: C.accentDark, background: C.accentSoft,
                        borderRadius: 999, padding: "1px 5px",
                      }}>
                        {initials(r.employeeId)}
                      </span>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
        <div style={{ display: "flex", gap: 16, marginTop: 14, fontFamily: sans, fontSize: 11.5, color: C.inkMuted }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <span style={{ width: 12, height: 12, borderRadius: 3, background: C.warnSoft, border: `1px solid ${C.border}` }} />
            Public holiday
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <span style={{ fontFamily: sans, fontSize: 9.5, fontWeight: 700, color: C.accentDark, background: C.accentSoft, borderRadius: 999, padding: "1px 5px" }}>AB</span>
            Approved time off (initials)
          </div>
        </div>
      </Panel>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Admin: Dashboard                                                        */
/* ---------------------------------------------------------------------- */

function AdminDashboard({ employees, clients, entries, rentalProperties, rentalExpenses }) {
  const [period, setPeriod] = useState("month"); // "week" | "month" | "quarter" | "year"
  const [anchor, setAnchor] = useState(TODAY);
  const DEFAULT_DASHBOARD_CONFIG = {
    statCards: true, growthPanel: true, allocationChart: true, revenueCostChart: true,
    byClientTable: true, byEmployeeTable: true, nonChargeableTable: true, profitabilityTable: true, departmentTable: true, costCenterTable: true,
    monthlyEmployeeTable: true, weeklyEmployeeTable: true, monthlyClientTable: true, companyMonthlyTable: true,
  };
  const [configRaw, setConfig] = useAppSetting("dashboard_config", DEFAULT_DASHBOARD_CONFIG);
  const config = configRaw || DEFAULT_DASHBOARD_CONFIG;
  const [customizing, setCustomizing] = useState(false);

  const [rangeStart, rangeEnd] = periodRange(period, anchor);

  const rangeEntries = entries.filter((e) => {
    const d = fromKey(e.date);
    return d >= rangeStart && d <= rangeEnd;
  });

  const activeEmployees = new Set(rangeEntries.map((e) => e.employeeId)).size;

  // Each employee's own hours and labor cost for the period, computed
  // once and reused below both to attribute cost to clients and to show
  // per-employee profitability.
  const employeeCosts = employees.filter((e) => employeeBilledInRange(e, rangeStart, rangeEnd)).map((emp) => {
    const empEntries = rangeEntries.filter((e) => e.employeeId === emp.id);
    return {
      id: emp.id,
      department: emp.department || DEFAULT_DEPARTMENT,
      entries: empEntries,
      hours: empEntries.reduce((s, e) => s + e.hours, 0),
      cost: periodEmployeeCost(emp, rangeStart, rangeEnd),
    };
  });

  // Active/inactive only controls whether a client can be picked for NEW
  // work (timesheet dropdown, new fees). For reporting, a client counts in
  // any period that overlaps when they were actually billed — governed by
  // their fee history start and their end date, not the active toggle —
  // so deactivating a client never erases its revenue from past periods.
  const byClient = clients.filter((c) => clientBilledInRange(c, rangeStart, rangeEnd)).map((c) => {
    const clientEntries = rangeEntries.filter((e) => e.clientId === c.id);
    const hrs = clientEntries.reduce((s, e) => s + e.hours, 0);
    const { revenue, allocation, revenueByCostCenter } = periodClientMetrics(c, clientEntries, rangeStart, rangeEnd);
    // Attribute each employee's labor cost to this client proportionally
    // to the hours they spent on it — the mirror image of how revenue is
    // attributed to employees below.
    const cost = employeeCosts.reduce((sum, ec) => {
      if (ec.hours <= 0) return sum;
      const empHrsForClient = clientEntries.filter((e) => e.employeeId === ec.id).reduce((s, e) => s + e.hours, 0);
      return sum + ec.cost * (empHrsForClient / ec.hours);
    }, 0);
    const profit = revenue - cost;
    const margin = revenue > 0 ? (profit / revenue) * 100 : null;
    return { ...c, hours: hrs, allocation, revenue, revenueByCostCenter, effectiveRate: hrs > 0 ? revenue / hrs : null, cost, profit, margin };
  }).sort((a, b) => b.hours - a.hours);

  const totalRevenue = byClient.reduce((s, c) => s + c.revenue, 0);
  const totalHours = rangeEntries.reduce((s, e) => s + e.hours, 0);
  const avgEffectiveRate = totalHours > 0 ? totalRevenue / totalHours : 0;

  // Attribute each client's revenue to employees proportionally to the
  // hours they logged for that client — a simple productivity signal,
  // not a payroll or commission calculation.
  const byEmployee = employees.filter((e) => employeeBilledInRange(e, rangeStart, rangeEnd)).map((emp) => {
    const hrs = rangeEntries.filter((e) => e.employeeId === emp.id).reduce((s, e) => s + e.hours, 0);
    const clientsTouched = new Set(rangeEntries.filter((e) => e.employeeId === emp.id && e.clientId).map((e) => e.clientId)).size;
    const nonChargeableHrs = rangeEntries.filter((e) => e.employeeId === emp.id && e.activity).reduce((s, e) => s + e.hours, 0);
    const revenue = byClient.reduce((sum, c) => {
      const clientTotalHrs = c.hours;
      if (clientTotalHrs <= 0) return sum;
      const empHrsForClient = rangeEntries.filter((e) => e.employeeId === emp.id && e.clientId === c.id).reduce((s, e) => s + e.hours, 0);
      return sum + c.revenue * (empHrsForClient / clientTotalHrs);
    }, 0);
    const cost = employeeCosts.find((ec) => ec.id === emp.id)?.cost || 0;
    const profit = revenue - cost;
    const margin = revenue > 0 ? (profit / revenue) * 100 : null;
    return { ...emp, hours: hrs, chargeableHours: hrs - nonChargeableHrs, nonChargeableHrs, clientsTouched, revenue, effectiveRate: hrs > 0 ? revenue / hrs : null, cost, profit, margin };
  }).sort((a, b) => b.hours - a.hours);

  // Non-chargeable time in the period, by activity — hours that carry no
  // revenue (business development, training, client communication, other).
  const totalHoursInPeriod = rangeEntries.reduce((s, e) => s + e.hours, 0);
  const nonChargeable = ACTIVITIES.map((a) => {
    const hours = rangeEntries.filter((e) => e.activity === a).reduce((s, e) => s + e.hours, 0);
    return { key: a, label: ACTIVITY_LABELS[a], hours, share: totalHoursInPeriod > 0 ? (hours / totalHoursInPeriod) * 100 : 0 };
  });
  const nonChargeableTotal = nonChargeable.reduce((s, r) => s + r.hours, 0);

  const totalLaborCost = byEmployee.reduce((s, e) => s + e.cost, 0);
  // Non-labor cost: rental expenses (rent, utilities, common expenses) booked
  // per property and month — all of it belongs to the Rental cost center.
  const rentalExp = rentalExpensesForRange(rentalExpenses, rangeStart, rangeEnd);
  const totalCost = totalLaborCost + rentalExp.total;
  const totalProfit = totalRevenue - totalCost;
  const overallMargin = totalRevenue > 0 ? (totalProfit / totalRevenue) * 100 : null;

  // Revenue per cost center first (needed as the allocation key for
  // departments that don't log hours), then each employee's cost spread
  // across their department's cost centers.
  const revenueByCC = emptyByCC();
  byClient.forEach((c) => { Object.keys(revenueByCC).forEach((cc) => { revenueByCC[cc] += c.revenueByCostCenter?.[cc] || 0; }); });
  employeeCosts.forEach((ec) => {
    const emp = employees.find((e) => e.id === ec.id);
    ec.byCostCenter = employeeCostByCostCenter(emp, ec.cost, ec.entries, revenueByCC);
  });
  const costByDepartment = Object.fromEntries(DEPARTMENTS.map((d) => [d.key, 0]));
  costByDepartment.rental += rentalExp.total;
  const unallocatedCostByDepartment = Object.fromEntries(DEPARTMENTS.map((d) => [d.key, 0]));
  employeeCosts.forEach((ec) => {
    costByDepartment[ec.department] = (costByDepartment[ec.department] || 0) + ec.cost;
    unallocatedCostByDepartment[ec.department] = (unallocatedCostByDepartment[ec.department] || 0) + (ec.byCostCenter[UNALLOCATED] || 0);
  });

  // Revenue, labor cost and profit per cost center for the period. The
  // rows add up to the same totals as the cards above; "Unallocated" only
  // appears when some fixed fee has no split yet or a department's cost had
  // no hours or revenue to spread it by.
  const byCostCenter = [...COST_CENTERS, UNALLOCATED].map((cc) => {
    const revenue = revenueByCC[cc] || 0;
    const cost = employeeCosts.reduce((s, ec) => s + (ec.byCostCenter?.[cc] || 0), 0) + (cc === "rental" ? rentalExp.total : 0);
    const profit = revenue - cost;
    const margin = revenue > 0 ? (profit / revenue) * 100 : null;
    return { key: cc, label: costCenterLabel(cc), revenue, cost, profit, margin };
  }).filter((r) => r.key !== UNALLOCATED || r.revenue > 0.005 || r.cost > 0.005);
  const costCenterChartData = byCostCenter.map((r) => ({ name: r.label, Revenue: Number(r.revenue.toFixed(2)), Cost: Number(r.cost.toFixed(2)) }));

  // Department P&L: revenue = its cost centers' revenue; cost = the labor
  // cost of the employees assigned to the department (including any part
  // not attributable to a single cost center). "Unallocated" revenue (fixed
  // fees with no split yet) is listed on its own so the total reconciles.
  const byDepartment = DEPARTMENTS.map((dep) => {
    const rows = byCostCenter.filter((r) => dep.costCenters.includes(r.key));
    const revenue = rows.reduce((s, r) => s + r.revenue, 0);
    const cost = costByDepartment[dep.key] || 0;
    const profit = revenue - cost;
    const headcount = employeeCosts.filter((ec) => ec.department === dep.key).length;
    return { key: dep.key, label: dep.label, costCenters: rows, unallocatedCost: unallocatedCostByDepartment[dep.key] || 0, headcount, revenue, cost, profit, margin: revenue > 0 ? (profit / revenue) * 100 : null };
  });
  // Compact profitability view: one line per department plus company total.
  const profitRows = [
    ...byDepartment.map((d) => ({ key: d.key, label: d.label, revenue: d.revenue, cost: d.cost, profit: d.profit, margin: d.margin })),
    ...((revenueByCC[UNALLOCATED] || 0) > 0.005 ? [{ key: UNALLOCATED, label: "Fees with no split yet", revenue: revenueByCC[UNALLOCATED], cost: 0, profit: revenueByCC[UNALLOCATED], margin: null }] : []),
    { key: "total", label: "Company total", revenue: totalRevenue, cost: totalCost, profit: totalProfit, margin: overallMargin },
  ];
  const unallocatedRevenue = revenueByCC[UNALLOCATED] || 0;
  const unallocatedRow = unallocatedRevenue > 0.005 ? { revenue: unallocatedRevenue, cost: 0, profit: unallocatedRevenue } : null;

  // Clients/employees that first started within THIS period specifically
  // (not just "existing by now") — driven by the same "Effective from" /
  // cost-history date you set, so it reflects what you actually told the
  // app, not when the database row happened to be created.
  const newClientsThisPeriod = clients.filter((c) => {
    const start = earliestHistoryDate(c.feeHistory);
    if (!start) return false;
    const d = fromKey(start);
    return d >= rangeStart && d <= rangeEnd;
  });
  const newEmployeesThisPeriod = employees.filter((e) => {
    const start = earliestHistoryDate(e.costHistory);
    if (!start) return false;
    const d = fromKey(start);
    return d >= rangeStart && d <= rangeEnd;
  });
  const newHiresCost = newEmployeesThisPeriod.reduce((sum, e) => {
    const ec = employeeCosts.find((x) => x.id === e.id);
    return sum + (ec ? ec.cost : 0);
  }, 0);

  // Rolling 6-month window (oldest to newest, ending on the anchor month)
  // for the per-employee, per-client, and company-wide monthly trends.
  // Each month uses whichever fee/cost terms were actually in effect.
  const trendMonths = Array.from({ length: 6 }, (_, i) => {
    const d = new Date(anchor.getFullYear(), anchor.getMonth() - (5 - i), 1);
    return { label: d.toLocaleDateString("en-GB", { month: "short", year: "2-digit" }), start: startOfMonth(d), end: endOfMonth(d) };
  });
  const employeeMonthlyHours = employees.filter((e) => employeeBilledInRange(e, trendMonths[0].start, trendMonths[trendMonths.length - 1].end)).map((emp) => ({
    ...emp,
    months: trendMonths.map((m) => entries.filter((e) => e.employeeId === emp.id && fromKey(e.date) >= m.start && fromKey(e.date) <= m.end).reduce((s, e) => s + e.hours, 0)),
  }));

  // Rolling 8-week window (oldest to newest, ending on the anchor's week)
  // for the per-employee weekly trend table.
  const trendWeeks = Array.from({ length: 8 }, (_, i) => {
    const start = addDays(startOfWeek(anchor), (i - 7) * 7);
    return { label: fmtShort(start), start, end: addDays(start, 6) };
  });
  const employeeWeeklyHours = employees.filter((e) => employeeBilledInRange(e, trendWeeks[0].start, trendWeeks[trendWeeks.length - 1].end)).map((emp) => ({
    ...emp,
    weeks: trendWeeks.map((w) => entries.filter((e) => e.employeeId === emp.id && fromKey(e.date) >= w.start && fromKey(e.date) <= w.end).reduce((s, e) => s + e.hours, 0)),
  }));

  const clientMonthly = clients.filter((c) => clientBilledInRange(c, trendMonths[0].start, trendMonths[trendMonths.length - 1].end)).map((c) => {
    const months = trendMonths.map((m) => {
      const monthEntries = entries.filter((e) => e.clientId === c.id && fromKey(e.date) >= m.start && fromKey(e.date) <= m.end);
      const hrs = monthEntries.reduce((s, e) => s + e.hours, 0);
      return { hours: hrs, revenue: periodClientMetrics(c, monthEntries, m.start, m.end).revenue };
    });
    return { ...c, months, totalRevenue: months.reduce((s, m) => s + m.revenue, 0) };
  });

  // Company-wide monthly overview: new clients added, revenue, cost, profit.
  const companyMonthly = trendMonths.map((m) => {
    // "New" = the month the client's revenue starts (earliest "Effective
    // from" in their fee history), not when the row was created.
    const newClients = clients.filter((c) => {
      const start = earliestHistoryDate(c.feeHistory);
      if (!start) return false;
      const d = fromKey(start);
      return d >= m.start && d <= m.end;
    }).length;
    const monthEntries = entries.filter((e) => fromKey(e.date) >= m.start && fromKey(e.date) <= m.end);
    const revenue = clients.filter((c) => clientBilledInRange(c, m.start, m.end)).reduce((sum, c) => sum + periodClientMetrics(c, monthEntries.filter((e) => e.clientId === c.id), m.start, m.end).revenue, 0);
    const cost = employees.filter((e) => employeeBilledInRange(e, m.start, m.end)).reduce((sum, e) => sum + periodEmployeeCost(e, m.start, m.end), 0)
      + rentalExpensesForRange(rentalExpenses, m.start, m.end).total;
    return { label: m.label, newClients, revenue, cost, profit: revenue - cost };
  });

  const chartData = byClient.filter((c) => c.hours > 0 || c.allocation).map((c) => ({
    name: c.name.length > 14 ? c.name.slice(0, 13) + "…" : c.name,
    Actual: Number(c.hours.toFixed(2)),
    Allocated: c.allocation ? Number(c.allocation.toFixed(2)) : 0,
  }));

  function shiftAnchor(dir) {
    setAnchor((a) => shiftPeriodAnchor(period, a, dir));
  }

  function exportReport() {
    exportToExcel(`dashboard-report-${toKey(TODAY)}.xlsx`, [
      { name: "By client", rows: byClient.map((c) => ({
        Client: c.name, Hours: Number(c.hours.toFixed(2)), "Allocation (h)": c.allocation ? Number(c.allocation.toFixed(1)) : "",
        "Revenue (€)": Number(c.revenue.toFixed(2)), "Cost (€)": Number(c.cost.toFixed(2)), "Profit (€)": Number(c.profit.toFixed(2)),
        "Margin (%)": c.margin !== null ? Number(c.margin.toFixed(1)) : "",
      })) },
      { name: "Non-chargeable time", rows: nonChargeable.map((r) => ({ Activity: r.label, Hours: Number(r.hours.toFixed(2)), "Share of all hours (%)": Number(r.share.toFixed(1)) })) },
      { name: "By employee", rows: byEmployee.map((e) => ({
        Employee: e.name, Hours: Number(e.hours.toFixed(2)), "Chargeable hours": Number(e.chargeableHours.toFixed(2)), "Non-chargeable hours": Number(e.nonChargeableHrs.toFixed(2)), Clients: e.clientsTouched,
        "Revenue (€)": Number(e.revenue.toFixed(2)), "Cost (€)": Number(e.cost.toFixed(2)), "Profit (€)": Number(e.profit.toFixed(2)),
        "Margin (%)": e.margin !== null ? Number(e.margin.toFixed(1)) : "",
      })) },
      { name: "Profitability", rows: profitRows.map((d) => ({
        Department: d.label, "Revenue (€)": Number(d.revenue.toFixed(2)), "Cost (€)": Number(d.cost.toFixed(2)),
        "Profit (€)": Number(d.profit.toFixed(2)), "Profit % of revenue": d.margin !== null ? Number(d.margin.toFixed(1)) : "",
      })) },
      { name: "By department", rows: [
        ...byDepartment.map((d) => ({
          Department: d.label, "Revenue (€)": Number(d.revenue.toFixed(2)), "Cost (€)": Number(d.cost.toFixed(2)),
          "Profit (€)": Number(d.profit.toFixed(2)), "Margin (%)": d.margin !== null ? Number(d.margin.toFixed(1)) : "",
        })),
        ...(unallocatedRow ? [{ Department: "Fees with no split yet", "Revenue (€)": Number(unallocatedRow.revenue.toFixed(2)), "Cost (€)": 0, "Profit (€)": Number(unallocatedRow.profit.toFixed(2)), "Margin (%)": "" }] : []),
        { Department: "Total", "Revenue (€)": Number(totalRevenue.toFixed(2)), "Cost (€)": Number(totalCost.toFixed(2)), "Profit (€)": Number(totalProfit.toFixed(2)), "Margin (%)": overallMargin !== null ? Number(overallMargin.toFixed(1)) : "" },
      ] },
      { name: "Rental expenses", rows: rentalProperties.map((pr) => {
        const x = rentalExp.byProperty[pr.id] || {};
        return { Property: pr.name, ...Object.fromEntries(RENTAL_EXPENSE_CATEGORIES.map((k) => [`${RENTAL_EXPENSE_LABELS[k]} (€)`, Number((x[k] || 0).toFixed(2))])), "Total (€)": Number((x.total || 0).toFixed(2)) };
      }) },
      { name: "By cost center", rows: byCostCenter.map((r) => ({
        Department: departmentOf(r.key)?.label || "Unallocated", "Cost center": r.label,
        "Revenue (€)": Number(r.revenue.toFixed(2)), "Cost (€)": Number(r.cost.toFixed(2)),
        "Profit (€)": Number(r.profit.toFixed(2)), "Margin (%)": r.margin !== null ? Number(r.margin.toFixed(1)) : "",
      })) },
      { name: "Company by month", rows: companyMonthly.map((m) => ({
        Month: m.label, "New clients": m.newClients, "Revenue (€)": Number(m.revenue.toFixed(2)),
        "Cost (€)": Number(m.cost.toFixed(2)), "Profit (€)": Number(m.profit.toFixed(2)),
      })) },
      { name: "Growth", rows: growthRows.map((r) => ({
        Metric: r.label, "This period": Number(r.value.toFixed(2)), "Previous period": Number(prevTotals[r.key].toFixed(2)),
        "vs. previous (%)": r.vsPrev !== null ? Number(r.vsPrev.toFixed(1)) : "", "Same period last year": Number(yoyTotals[r.key].toFixed(2)),
        "YoY growth (%)": r.vsYoy !== null ? Number(r.vsYoy.toFixed(1)) : "",
      })) },
    ]);
  }

  // Growth: this period vs. the immediately preceding one of the same
  // length, and vs. the same calendar period a year earlier (YoY).
  const { prevStart, prevEnd, yoyStart, yoyEnd } = comparisonRanges(rangeStart, rangeEnd);
  const prevTotals = computePeriodTotals(clients, employees, entries, rentalExpenses, prevStart, prevEnd);
  const yoyTotals = computePeriodTotals(clients, employees, entries, rentalExpenses, yoyStart, yoyEnd);
  const growthRows = [
    { key: "revenue", label: "Revenue", value: totalRevenue, fmt: fmtEur },
    { key: "cost", label: "Cost", value: totalCost, fmt: fmtEur },
    { key: "profit", label: "Profit", value: totalProfit, fmt: fmtEur },
    { key: "hours", label: "Hours", value: totalHours, fmt: (v) => v.toFixed(2) + "h" },
  ].map((r) => ({ ...r, vsPrev: pctChange(r.value, prevTotals[r.key]), vsYoy: pctChange(r.value, yoyTotals[r.key]) }));

  const CONFIG_LABELS = {
    statCards: "Summary cards", growthPanel: "Growth — vs. previous period and YoY", allocationChart: "Actual vs. allocated chart", revenueCostChart: "Revenue vs. cost chart",
    byClientTable: "Hours by client table", byEmployeeTable: "Hours by employee table", nonChargeableTable: "Non-chargeable time", profitabilityTable: "Profitability by department", departmentTable: "P&L by department (detail)", costCenterTable: "Revenue & cost by cost center",
    monthlyEmployeeTable: "Hours per employee, by month", weeklyEmployeeTable: "Hours per employee, by week", monthlyClientTable: "Client analysis, by month",
    companyMonthlyTable: "Company overview, by month",
  };

  return (
    <div>
      <PageHeader
        title="Dashboard"
        sub="Hours across the company, by employee and by client."
        right={
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            <div style={{ display: "flex", border: `1px solid ${C.border}`, borderRadius: 7, overflow: "hidden" }}>
              {PERIOD_OPTIONS.map((p) => (
                <button key={p.key} onClick={() => setPeriod(p.key)}
                  style={{
                    padding: "7px 11px", fontSize: 12.5, fontFamily: sans, fontWeight: 600, border: "none", cursor: "pointer",
                    background: period === p.key ? C.accent : C.surface, color: period === p.key ? "#fff" : C.inkMuted,
                  }}>
                  {p.label}
                </button>
              ))}
            </div>
            <button onClick={() => shiftAnchor(-1)} style={navBtnStyle}><ChevronLeft size={16} /></button>
            <div style={{ fontFamily: sans, fontSize: 13, fontWeight: 600, minWidth: 150, textAlign: "center" }}>
              {periodLabel(period, rangeStart, rangeEnd)}
            </div>
            <button onClick={() => shiftAnchor(1)} style={navBtnStyle}><ChevronRight size={16} /></button>
            <Btn variant="secondary" size="sm" onClick={() => setCustomizing(true)}>Customize</Btn>
            <Btn variant="secondary" size="sm" onClick={exportReport}>Export to Excel</Btn>
          </div>
        }
      />

      {config.statCards && (
        <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
          <StatCard label="Active employees" value={activeEmployees} sub={`of ${employees.filter(e=>e.role==='employee' && (!earliestHistoryDate(e.costHistory) || earliestHistoryDate(e.costHistory) <= toKey(rangeEnd))).length} employees`} />
          <StatCard label="Clients with hours" value={byClient.filter((c) => c.hours > 0).length} sub={`of ${byClient.length} active clients`} />
          <StatCard label="New clients this period" value={newClientsThisPeriod.length} sub={newClientsThisPeriod.length ? newClientsThisPeriod.map(c=>c.name).join(", ") : "None started this period"} />
          <StatCard label="New employees this period" value={newEmployeesThisPeriod.length} sub={newEmployeesThisPeriod.length ? fmtEur(newHiresCost) + " added to labor cost" : "None started this period"} />
          <StatCard label="Revenue this period" value={fmtEur(totalRevenue)} accent={C.accent} sub="Fixed fees + extra fees" />
          <StatCard label="Labor cost this period" value={fmtEur(totalLaborCost)} sub="Gross + benefits, prorated" />
          <StatCard label="Rental expenses this period" value={fmtEur(rentalExp.total)} sub="Rent + utilities + common expenses" />
          <StatCard label="Profit this period" value={fmtEur(totalProfit)} accent={totalProfit < 0 ? C.danger : C.accent} sub="Revenue − labor − rental expenses" />
          <StatCard label="Overall margin" value={overallMargin !== null ? overallMargin.toFixed(0) + "%" : "—"} accent={overallMargin !== null && overallMargin < 0 ? C.danger : C.accent} sub="(Revenue − cost) ÷ revenue" />
          <StatCard label="Avg. effective rate" value={fmtEur(avgEffectiveRate) + "/h"} sub="Revenue ÷ hours worked" />
        </div>
      )}

      {config.growthPanel && (
        <Panel style={{ marginBottom: 18 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Growth</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>
            This period ({fmtShort(rangeStart)} – {fmtShort(rangeEnd)}) vs. the previous period of the same length ({fmtShort(prevStart)} – {fmtShort(prevEnd)}) and the same period a year earlier ({fmtShort(yoyStart)} – {fmtShort(yoyEnd)}).
          </div>
          <TableShell headers={["Metric", "This period", "Previous period", "vs. previous", "Same period last year", "YoY growth"]}>
            {growthRows.map((r) => {
              const growthTone = (v) => v === null ? undefined : v < 0 ? C.danger : v > 0 ? C.accentDark : C.inkMuted;
              const growthText = (v) => v === null ? "—" : `${v > 0 ? "+" : ""}${v.toFixed(1)}%`;
              return (
                <tr key={r.key}>
                  <Td>{r.label}</Td>
                  <Td mono style={{ fontWeight: 700 }}>{r.fmt(r.value)}</Td>
                  <Td mono style={{ color: C.inkMuted }}>{r.fmt(prevTotals[r.key])}</Td>
                  <Td mono style={{ color: growthTone(r.vsPrev), fontWeight: 700 }}>{growthText(r.vsPrev)}</Td>
                  <Td mono style={{ color: C.inkMuted }}>{r.fmt(yoyTotals[r.key])}</Td>
                  <Td mono style={{ color: growthTone(r.vsYoy), fontWeight: 700 }}>{growthText(r.vsYoy)}</Td>
                </tr>
              );
            })}
          </TableShell>
        </Panel>
      )}

      {config.allocationChart && (
        <Panel style={{ marginBottom: 18 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 14 }}>
            Actual vs. allocated hours per client
          </div>
          {chartData.length === 0 ? (
            <EmptyState icon={Building2} title="No hours in this period" />
          ) : (
            <div style={{ width: "100%", height: 260 }}>
              <ResponsiveContainer>
                <BarChart data={chartData} margin={{ left: -10, right: 10 }}>
                  <CartesianGrid stroke={C.border} vertical={false} />
                  <XAxis dataKey="name" tick={{ fontFamily: sans, fontSize: 11, fill: C.inkMuted }} axisLine={{ stroke: C.border }} tickLine={false} />
                  <YAxis tick={{ fontFamily: mono, fontSize: 11, fill: C.inkMuted }} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ fontFamily: sans, fontSize: 12.5, borderRadius: 8, border: `1px solid ${C.border}` }} />
                  <Bar dataKey="Actual" fill={C.accent} radius={[4, 4, 0, 0]} maxBarSize={26} />
                  <Bar dataKey="Allocated" fill={C.border} radius={[4, 4, 0, 0]} maxBarSize={26} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
        </Panel>
      )}

      {config.revenueCostChart && (
        <Panel style={{ marginBottom: 18 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Revenue vs. cost, by month</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 14 }}>Last 6 months, ending {trendMonths[5].label}.</div>
          <div style={{ width: "100%", height: 260 }}>
            <ResponsiveContainer>
              <LineChart data={companyMonthly} margin={{ left: -10, right: 10 }}>
                <CartesianGrid stroke={C.border} vertical={false} />
                <XAxis dataKey="label" tick={{ fontFamily: sans, fontSize: 11, fill: C.inkMuted }} axisLine={{ stroke: C.border }} tickLine={false} />
                <YAxis tick={{ fontFamily: mono, fontSize: 11, fill: C.inkMuted }} axisLine={false} tickLine={false} />
                <Tooltip contentStyle={{ fontFamily: sans, fontSize: 12.5, borderRadius: 8, border: `1px solid ${C.border}` }} formatter={(v) => fmtEur(v)} />
                <Legend wrapperStyle={{ fontFamily: sans, fontSize: 12 }} />
                <Line type="monotone" dataKey="revenue" name="Revenue" stroke={C.accent} strokeWidth={2.5} dot={{ r: 3 }} />
                <Line type="monotone" dataKey="cost" name="Cost" stroke={C.danger} strokeWidth={2.5} dot={{ r: 3 }} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </Panel>
      )}

      {config.byClientTable && (
        <Panel style={{ marginBottom: 18 }}>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 12 }}>Hours by client</div>
            <TableShell headers={["Client", "Hours", "Allocation", "Status", "Revenue", "Cost", "Profit", "Margin"]}>
              {byClient.map((c) => {
                const pct = c.allocation ? (c.hours / c.allocation) * 100 : null;
                let tone = "neutral", label = "No allocation set";
                if (pct !== null) {
                  if (pct > 100) { tone = "danger"; label = `${pct.toFixed(0)}% — over`; }
                  else if (pct >= 85) { tone = "warn"; label = `${pct.toFixed(0)}% — near limit`; }
                  else { tone = "accent"; label = `${pct.toFixed(0)}% used`; }
                }
                return (
                  <tr key={c.id}>
                    <Td>{c.name}</Td>
                    <Td mono>{c.hours.toFixed(2)}h</Td>
                    <Td mono>{c.allocation ? c.allocation.toFixed(1) + "h" : "—"}</Td>
                    <Td><Badge tone={tone}>{label}</Badge></Td>
                    <Td mono>{fmtEur(c.revenue)}</Td>
                    <Td mono>{fmtEur(c.cost)}</Td>
                    <Td mono style={{ color: c.profit < 0 ? C.danger : undefined }}>{fmtEur(c.profit)}</Td>
                    <Td>{c.margin !== null ? <Badge tone={c.margin < 0 ? "danger" : c.margin < 20 ? "warn" : "accent"}>{c.margin.toFixed(0)}%</Badge> : "—"}</Td>
                  </tr>
                );
              })}
            </TableShell>
        </Panel>
      )}

      {config.byEmployeeTable && (
        <Panel style={{ marginBottom: 18 }}>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 12 }}>Hours by employee</div>
            <TableShell headers={["Employee", "Hours", "Chargeable", "Non-chargeable", "Clients", "Revenue", "Cost", "Profit", "Margin"]}>
              {byEmployee.map((e) => (
                <tr key={e.id}>
                  <Td>{e.name}</Td>
                  <Td mono>{e.hours.toFixed(2)}h</Td>
                  <Td mono>{e.chargeableHours.toFixed(2)}h{e.hours > 0 ? <span style={{ color: C.inkFaint, fontWeight: 400 }}> ({(e.chargeableHours / e.hours * 100).toFixed(0)}%)</span> : null}</Td>
                  <Td mono style={{ color: e.nonChargeableHrs > 0 ? C.warn : undefined }}>{e.nonChargeableHrs > 0 ? e.nonChargeableHrs.toFixed(2) + "h" : "—"}</Td>
                  <Td mono>{e.clientsTouched}</Td>
                  <Td mono>{fmtEur(e.revenue)}</Td>
                  <Td mono>{fmtEur(e.cost)}</Td>
                  <Td mono style={{ color: e.profit < 0 ? C.danger : undefined }}>{fmtEur(e.profit)}</Td>
                  <Td>{e.margin !== null ? <Badge tone={e.margin < 0 ? "danger" : e.margin < 20 ? "warn" : "accent"}>{e.margin.toFixed(0)}%</Badge> : "—"}</Td>
                </tr>
              ))}
            </TableShell>
        </Panel>
      )}

      {config.nonChargeableTable && (
        <Panel style={{ marginBottom: 18 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Non-chargeable time</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>
            Hours logged to internal activities instead of a client — {nonChargeableTotal.toFixed(2)}h of {totalHoursInPeriod.toFixed(2)}h in the period ({totalHoursInPeriod > 0 ? (nonChargeableTotal / totalHoursInPeriod * 100).toFixed(0) : 0}%). These hours carry no revenue and are not attributed to clients.
          </div>
          <TableShell headers={["Activity", "Hours", "Share of all hours"]}>
            {nonChargeable.map((r) => (
              <tr key={r.key}>
                <Td>{r.label}</Td>
                <Td mono>{r.hours.toFixed(2)}h</Td>
                <Td mono style={{ color: C.inkMuted }}>{r.share.toFixed(1)}%</Td>
              </tr>
            ))}
            <tr style={{ fontWeight: 700 }}>
              <Td>Total non-chargeable</Td>
              <Td mono>{nonChargeableTotal.toFixed(2)}h</Td>
              <Td mono>{totalHoursInPeriod > 0 ? (nonChargeableTotal / totalHoursInPeriod * 100).toFixed(1) : "0.0"}%</Td>
            </tr>
          </TableShell>
        </Panel>
      )}

      {config.profitabilityTable && (
        <Panel style={{ marginBottom: 18 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Profitability by department</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Revenue − Cost = Profit, and Profit as a % of Revenue, for {periodLabel(period, rangeStart, rangeEnd)}.</div>
          <TableShell headers={["Department", "Revenue", "− Cost", "= Profit", "Profit % of revenue", "Share of company profit"]}>
            {profitRows.map((d) => (
              <tr key={d.key} style={{ fontWeight: d.key === "total" ? 700 : 400, borderTop: d.key === "total" ? `2px solid ${C.borderStrong}` : undefined }}>
                <Td>{d.label}</Td>
                <Td mono>{fmtEur(d.revenue)}</Td>
                <Td mono>{fmtEur(d.cost)}</Td>
                <Td mono style={{ color: d.profit < 0 ? C.danger : C.accentDark, fontWeight: 700 }}>{fmtEur(d.profit)}</Td>
                <Td>{d.margin !== null ? <Badge tone={d.margin < 0 ? "danger" : d.margin < 20 ? "warn" : "accent"}>{d.margin.toFixed(1)}%</Badge> : "—"}</Td>
                <Td mono style={{ color: C.inkMuted }}>{d.key === "total" ? "100%" : (totalProfit > 0 ? (d.profit / totalProfit * 100).toFixed(0) + "%" : "—")}</Td>
              </tr>
            ))}
          </TableShell>
        </Panel>
      )}

      {config.departmentTable && (
        <Panel style={{ marginBottom: 18 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>P&amp;L by department</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>
            {DEPARTMENTS.map((d) => `${d.label} = ${d.costCenters.map((cc) => COST_CENTER_LABELS[cc]).join(" / ")}`).join(" · ")}
          </div>
          <TableShell headers={["Department", "People", "Revenue", "Cost", "Profit", "Margin"]}>
            {byDepartment.map((d) => (
              <React.Fragment key={d.key}>
                <tr style={{ fontWeight: 700, background: C.accentSoft }}>
                  <Td>{d.label}</Td>
                  <Td mono>{d.headcount}</Td>
                  <Td mono>{fmtEur(d.revenue)}</Td>
                  <Td mono>{fmtEur(d.cost)}</Td>
                  <Td mono style={{ color: d.profit < 0 ? C.danger : undefined }}>{fmtEur(d.profit)}</Td>
                  <Td>{d.margin !== null ? <Badge tone={d.margin < 0 ? "danger" : d.margin < 20 ? "warn" : "accent"}>{d.margin.toFixed(0)}%</Badge> : "—"}</Td>
                </tr>
                {d.costCenters.map((r) => (
                  <tr key={r.key}>
                    <Td style={{ paddingLeft: 28, color: C.inkMuted }}>{r.label}</Td>
                    <Td />
                    <Td mono style={{ color: C.inkMuted }}>{fmtEur(r.revenue)}</Td>
                    <Td mono style={{ color: C.inkMuted }}>{fmtEur(r.cost)}</Td>
                    <Td mono style={{ color: r.profit < 0 ? C.danger : C.inkMuted }}>{fmtEur(r.profit)}</Td>
                    <Td style={{ color: C.inkMuted }}>{r.margin !== null ? r.margin.toFixed(0) + "%" : "—"}</Td>
                  </tr>
                ))}
                {d.unallocatedCost > 0.005 && (
                  <tr>
                    <Td style={{ paddingLeft: 28, color: C.warn }}>Not attributable to a cost center</Td>
                    <Td />
                    <Td mono style={{ color: C.inkMuted }}>—</Td>
                    <Td mono style={{ color: C.warn }}>{fmtEur(d.unallocatedCost)}</Td>
                    <Td mono style={{ color: C.inkMuted }}>—</Td>
                    <Td />
                  </tr>
                )}
              </React.Fragment>
            ))}
            {unallocatedRow && (
              <tr style={{ opacity: 0.7 }}>
                <Td><Badge tone="warn">Fees with no split yet</Badge></Td>
                <Td />
                <Td mono>{fmtEur(unallocatedRow.revenue)}</Td>
                <Td mono>{fmtEur(unallocatedRow.cost)}</Td>
                <Td mono style={{ color: unallocatedRow.profit < 0 ? C.danger : undefined }}>{fmtEur(unallocatedRow.profit)}</Td>
                <Td>—</Td>
              </tr>
            )}
            <tr style={{ fontWeight: 700, borderTop: `2px solid ${C.borderStrong}` }}>
              <Td>Company total</Td>
              <Td mono>{employeeCosts.length}</Td>
              <Td mono>{fmtEur(totalRevenue)}</Td>
              <Td mono>{fmtEur(totalCost)}</Td>
              <Td mono style={{ color: totalProfit < 0 ? C.danger : undefined }}>{fmtEur(totalProfit)}</Td>
              <Td>{overallMargin !== null ? overallMargin.toFixed(0) + "%" : "—"}</Td>
            </tr>
          </TableShell>
        </Panel>
      )}

      {config.costCenterTable && (
        <Panel style={{ marginBottom: 18 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Revenue &amp; cost by cost center</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>
            Fixed fees follow each client's cost-center split, extra fees their own tag. Labor cost follows each employee's department: Accounting Team by logged hours, other departments in proportion to the department's revenue by cost center. Rental expenses (rent, utilities, common expenses) are added to the Rental cost center.
          </div>
          <div style={{ width: "100%", height: 220, marginBottom: 14 }}>
            <ResponsiveContainer>
              <BarChart data={costCenterChartData} margin={{ left: -10, right: 10 }}>
                <CartesianGrid stroke={C.border} vertical={false} />
                <XAxis dataKey="name" tick={{ fontFamily: sans, fontSize: 11, fill: C.inkMuted }} axisLine={{ stroke: C.border }} tickLine={false} />
                <YAxis tick={{ fontFamily: mono, fontSize: 11, fill: C.inkMuted }} axisLine={false} tickLine={false} />
                <Tooltip contentStyle={{ fontFamily: sans, fontSize: 12.5, borderRadius: 8, border: `1px solid ${C.border}` }} formatter={(v) => fmtEur(v)} />
                <Legend wrapperStyle={{ fontFamily: sans, fontSize: 12 }} />
                <Bar dataKey="Revenue" fill={C.accent} radius={[4, 4, 0, 0]} maxBarSize={26} />
                <Bar dataKey="Cost" fill={C.danger} radius={[4, 4, 0, 0]} maxBarSize={26} />
              </BarChart>
            </ResponsiveContainer>
          </div>
          <TableShell headers={["Cost center", "Revenue", "Cost", "Profit", "Margin"]}>
            {byCostCenter.map((r) => (
              <tr key={r.key} style={{ opacity: r.key === UNALLOCATED ? 0.7 : 1 }}>
                <Td>{r.key === UNALLOCATED ? <Badge tone="warn">{r.label}</Badge> : r.label}</Td>
                <Td mono>{fmtEur(r.revenue)}</Td>
                <Td mono>{fmtEur(r.cost)}</Td>
                <Td mono style={{ color: r.profit < 0 ? C.danger : undefined }}>{fmtEur(r.profit)}</Td>
                <Td>{r.margin !== null ? <Badge tone={r.margin < 0 ? "danger" : r.margin < 20 ? "warn" : "accent"}>{r.margin.toFixed(0)}%</Badge> : "—"}</Td>
              </tr>
            ))}
            <tr style={{ fontWeight: 700 }}>
              <Td>Total</Td>
              <Td mono>{fmtEur(totalRevenue)}</Td>
              <Td mono>{fmtEur(totalCost)}</Td>
              <Td mono style={{ color: totalProfit < 0 ? C.danger : undefined }}>{fmtEur(totalProfit)}</Td>
              <Td>{overallMargin !== null ? overallMargin.toFixed(0) + "%" : "—"}</Td>
            </tr>
          </TableShell>
        </Panel>
      )}

      {config.companyMonthlyTable && (
        <Panel style={{ marginBottom: 18 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Company overview, by month</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>New clients, revenue, cost (labor + rental expenses), and profit — last 6 months.</div>
          <TableShell headers={["Month", "New clients", "Revenue", "Cost", "Profit"]}>
            {companyMonthly.map((m, i) => (
              <tr key={i}>
                <Td>{m.label}</Td>
                <Td mono>{m.newClients}</Td>
                <Td mono>{fmtEur(m.revenue)}</Td>
                <Td mono>{fmtEur(m.cost)}</Td>
                <Td mono style={{ color: m.profit < 0 ? C.danger : undefined }}>{fmtEur(m.profit)}</Td>
              </tr>
            ))}
          </TableShell>
        </Panel>
      )}

      <div style={{ display: "flex", gap: 18, flexWrap: "wrap", marginTop: 18 }}>
        {config.monthlyEmployeeTable && (
          <Panel style={{ flex: 1, minWidth: 380 }}>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Hours per employee, by month</div>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Last 6 months, ending {trendMonths[5].label}.</div>
            <TableShell headers={["Employee", ...trendMonths.map((m) => m.label), "Total"]}>
              {employeeMonthlyHours.map((e) => (
                <tr key={e.id}>
                  <Td>{e.name}</Td>
                  {e.months.map((h, i) => <Td key={i} mono>{h > 0 ? h.toFixed(1) : "—"}</Td>)}
                  <Td mono>{e.months.reduce((s, h) => s + h, 0).toFixed(1)}</Td>
                </tr>
              ))}
            </TableShell>
          </Panel>
        )}

        {config.monthlyClientTable && (
          <Panel style={{ flex: 1, minWidth: 380 }}>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Client analysis, by month</div>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Hours per month, plus total revenue over the window.</div>
            <TableShell headers={["Client", ...trendMonths.map((m) => m.label), "Revenue (6mo)"]}>
              {clientMonthly.map((c) => (
                <tr key={c.id}>
                  <Td>{c.name}</Td>
                  {c.months.map((m, i) => <Td key={i} mono>{m.hours > 0 ? m.hours.toFixed(1) : "—"}</Td>)}
                  <Td mono>{fmtEur(c.totalRevenue)}</Td>
                </tr>
              ))}
            </TableShell>
          </Panel>
        )}
      </div>

      {config.weeklyEmployeeTable && (
        <Panel style={{ marginTop: 18 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Hours per employee, by week</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Last 8 weeks (starting date shown), ending the week of {trendWeeks[7].label}.</div>
          <TableShell headers={["Employee", ...trendWeeks.map((w) => w.label), "Total"]}>
            {employeeWeeklyHours.map((e) => (
              <tr key={e.id}>
                <Td>{e.name}</Td>
                {e.weeks.map((h, i) => <Td key={i} mono>{h > 0 ? h.toFixed(1) : "—"}</Td>)}
                <Td mono>{e.weeks.reduce((s, h) => s + h, 0).toFixed(1)}</Td>
              </tr>
            ))}
          </TableShell>
        </Panel>
      )}

      {customizing && (
        <Modal title="Customize dashboard" onClose={() => setCustomizing(false)} width={380}>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted, marginBottom: 4 }}>
              Choose which sections to show. This applies for everyone using the dashboard.
            </div>
            {Object.keys(CONFIG_LABELS).map((key) => (
              <label key={key} style={{ display: "flex", alignItems: "center", gap: 10, fontFamily: sans, fontSize: 13.5, color: C.ink, cursor: "pointer" }}>
                <input type="checkbox" checked={!!config[key]} onChange={(e) => setConfig({ ...config, [key]: e.target.checked })} />
                {CONFIG_LABELS[key]}
              </label>
            ))}
            <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 10 }}>
              <Btn onClick={() => setCustomizing(false)}>Done</Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}



/* ---------------------------------------------------------------------- */
/* Admin: All timesheets (view + edit any entry)                          */
/* ---------------------------------------------------------------------- */

function AdminTimesheets({ employees, clients, entries, refetchEntries, lockedWeeks, refetchLockedWeeks }) {
  const [weekStart, setWeekStart] = useState(startOfWeek(TODAY));
  const [empFilter, setEmpFilter] = useState("all");
  const [clientFilter, setClientFilter] = useState("all");
  const [sortBy, setSortBy] = useState("date_asc");
  const [editingId, setEditingId] = useState(null);
  const [editHours, setEditHours] = useState("");
  const [editNote, setEditNote] = useState("");

  const weekEnd = addDays(weekStart, 6);
  const empName = (id) => employees.find((e) => e.id === id)?.name || "—";

  const SORTERS = {
    date_asc: (a, b) => a.date.localeCompare(b.date) || empName(a.employeeId).localeCompare(empName(b.employeeId)),
    date_desc: (a, b) => b.date.localeCompare(a.date) || empName(a.employeeId).localeCompare(empName(b.employeeId)),
    employee: (a, b) => empName(a.employeeId).localeCompare(empName(b.employeeId)) || a.date.localeCompare(b.date),
    hours_desc: (a, b) => b.hours - a.hours || a.date.localeCompare(b.date),
  };
  const rows = entries
    .filter((e) => { const d = fromKey(e.date); return d >= weekStart && d <= weekEnd; })
    .filter((e) => empFilter === "all" || e.employeeId === empFilter)
    .filter((e) => clientFilter === "all" || (clientFilter === "internal" ? !!e.activity : e.clientId === clientFilter))
    .sort(SORTERS[sortBy] || SORTERS.date_asc);

  async function saveEdit(id) {
    const val = Number(editHours);
    if (val > 0) { await supabase.from("time_entries").update({ hours: val, note: editNote.trim() || null }).eq("id", id); refetchEntries(); }
    setEditingId(null);
  }
  async function removeEntry(id) { await supabase.from("time_entries").delete().eq("id", id); refetchEntries(); }
  function isRowLocked(e) { return lockedWeeks.includes(`${e.employeeId}|${toKey(startOfWeek(fromKey(e.date)))}`); }
  async function toggleRowLock(e) {
    const wk = toKey(startOfWeek(fromKey(e.date)));
    if (isRowLocked(e)) await supabase.from("locked_weeks").delete().eq("employee_id", e.employeeId).eq("week_start", wk);
    else await supabase.from("locked_weeks").insert({ employee_id: e.employeeId, week_start: wk });
    refetchLockedWeeks();
  }

  const total = rows.reduce((s, e) => s + e.hours, 0);

  const weekEntries = entries.filter((e) => { const d = fromKey(e.date); return d >= weekStart && d <= weekEnd; });
  const compliance = employees.filter((e) => employeeBilledInRange(e, weekStart, weekEnd)).map((e) => {
    const target = e.weeklyHours || 40;
    const logged = weekEntries.filter((t) => t.employeeId === e.id).reduce((s, t) => s + t.hours, 0);
    const diff = logged - target;
    return { ...e, target, logged, diff };
  }).sort((a, b) => b.diff - a.diff);

  return (
    <div>
      <PageHeader title="All Timesheets" sub="Review and correct any employee's logged hours." />

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 4 }}>
          Weekly hours vs. contract
        </div>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>
          {fmtShort(weekStart)} – {fmtShort(weekEnd)}. Each employee's contracted hours are set on the Employees page.
        </div>
        <TableShell headers={["Employee", "Contracted /wk", "Logged", "Difference", "Status"]}>
          {compliance.map((e) => {
            const onTrack = Math.abs(e.diff) < 0.01;
            const over = e.diff > 0;
            const tone = onTrack ? "accent" : over ? "warn" : "danger";
            const label = onTrack ? "On track" : over ? `+${e.diff.toFixed(2)}h overtime` : `${e.diff.toFixed(2)}h undertime`;
            return (
              <tr key={e.id}>
                <Td>{e.name}</Td>
                <Td mono>{e.target}h</Td>
                <Td mono>{e.logged.toFixed(2)}h</Td>
                <Td mono>{onTrack ? "—" : `${over ? "+" : ""}${e.diff.toFixed(2)}h`}</Td>
                <Td><Badge tone={tone}>{label}</Badge></Td>
              </tr>
            );
          })}
        </TableShell>
      </Panel>

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end", marginBottom: 16 }}>
          <WeekNav weekStart={weekStart} setWeekStart={setWeekStart} />
          <Field label="Employee">
            <select style={{ ...inputStyle, width: 170 }} value={empFilter} onChange={(e) => setEmpFilter(e.target.value)}>
              <option value="all">All employees</option>
              {employees.filter((e) => e.active).map((e) => <option key={e.id} value={e.id}>{e.name}</option>)}
            </select>
          </Field>
          <Field label="Client">
            <select style={{ ...inputStyle, width: 170 }} value={clientFilter} onChange={(e) => setClientFilter(e.target.value)}>
              <option value="all">All clients &amp; activities</option>
              <option value="internal">Internal — non-chargeable only</option>
              {clients.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
          <Field label="Sort by">
            <select style={{ ...inputStyle, width: 170 }} value={sortBy} onChange={(e) => setSortBy(e.target.value)}>
              <option value="date_asc">Day (earliest first)</option>
              <option value="date_desc">Day (latest first)</option>
              <option value="employee">Employee name</option>
              <option value="hours_desc">Hours (most first)</option>
            </select>
          </Field>
        </div>

        {rows.length === 0 ? (
          <EmptyState icon={Clock} title="No entries match these filters" />
        ) : (
          <TableShell headers={["Date", "Employee", "Client / activity", "Category", "Hours", "Note", "Status", ""]}>
            {rows.map((e) => {
              const locked = isRowLocked(e);
              return (
                <tr key={e.id}>
                  <Td>{fmtDow(fromKey(e.date))} {fmtShort(fromKey(e.date))}</Td>
                  <Td>{empName(e.employeeId)}</Td>
                  <Td>{entryTargetLabel(e, clients)}</Td>
                  <Td>{e.activity ? <Badge tone="warn">Non-chargeable</Badge> : (CATEGORY_LABELS[e.category] || "Other")}</Td>
                  <Td mono>
                    {editingId === e.id ? (
                      <input type="number" min="0.25" step="0.25" autoFocus value={editHours}
                        onChange={(ev) => setEditHours(ev.target.value)}
                        style={{ ...inputStyle, width: 70, padding: "3px 7px", fontFamily: mono }} />
                    ) : e.hours.toFixed(2) + "h"}
                  </Td>
                  <Td>
                    {editingId === e.id ? (
                      <input type="text" placeholder="Note" value={editNote}
                        onChange={(ev) => setEditNote(ev.target.value)}
                        style={{ ...inputStyle, width: 160, padding: "3px 7px" }} />
                    ) : (e.note || "—")}
                  </Td>
                  <Td>{locked ? <Badge tone="warn">Submitted</Badge> : <Badge tone="neutral">Open</Badge>}</Td>
                  <Td right>
                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      {editingId === e.id ? (
                        <>
                          <button onClick={() => saveEdit(e.id)} style={iconBtnStyle}><Check size={14} color={C.accent} /></button>
                          <button onClick={() => setEditingId(null)} style={iconBtnStyle}><X size={14} color={C.inkMuted} /></button>
                        </>
                      ) : (
                        <>
                          <button onClick={() => { setEditingId(e.id); setEditHours(String(e.hours)); setEditNote(e.note || ""); }} style={iconBtnStyle}>
                            <Pencil size={13} color={C.inkMuted} />
                          </button>
                          <button onClick={() => removeEntry(e.id)} style={iconBtnStyle}>
                            <Trash2 size={13} color={C.inkMuted} />
                          </button>
                          <button onClick={() => toggleRowLock(e)} style={iconBtnStyle} title={locked ? "Reopen week" : "Lock week"}>
                            {locked ? <Unlock size={13} color={C.inkMuted} /> : <Lock size={13} color={C.inkMuted} />}
                          </button>
                        </>
                      )}
                    </div>
                  </Td>
                </tr>
              );
            })}
          </TableShell>
        )}
        <div style={{ display: "flex", justifyContent: "flex-end", paddingTop: 12, fontFamily: sans, fontSize: 13, fontWeight: 700 }}>
          Total:&nbsp;<span style={{ fontFamily: mono, color: C.accent }}>{total.toFixed(2)}h</span>
        </div>
      </Panel>
    </div>
  );
}


/* ---------------------------------------------------------------------- */
/* Admin: Clients                                                          */
/* ---------------------------------------------------------------------- */

function AdminClients({ clients, refetchClients, entries }) {
  const [modal, setModal] = useState(null); // null | "new" | client object
  const [name, setName] = useState("");
  const [allocAccounting, setAllocAccounting] = useState("");
  const [allocTax, setAllocTax] = useState("");
  const [allocPayroll, setAllocPayroll] = useState("");
  const [allocOther, setAllocOther] = useState("");
  const [feeSplit, setFeeSplit] = useState(() => Object.fromEntries(COST_CENTERS.map((cc) => [cc, ""])));
  const [feeSplitNote, setFeeSplitNote] = useState("");
  const [fixedFee, setFixedFee] = useState("");
  const [feeEffectiveDate, setFeeEffectiveDate] = useState(toKey(TODAY));
  const [endDate, setEndDate] = useState("");
  const [query, setQuery] = useState("");
  const [editingHistoryId, setEditingHistoryId] = useState(null);
  const [histForm, setHistForm] = useState(null);
  const [viewPeriod, setViewPeriod] = useState("month");
  const [viewAnchor, setViewAnchor] = useState(TODAY);
  const [extraFeeClientId, setExtraFeeClientId] = useState(null);
  const [extraFeeEditingId, setExtraFeeEditingId] = useState(null);
  const [extraFeeMonth, setExtraFeeMonth] = useState("");
  const [extraFeeAmount, setExtraFeeAmount] = useState("");
  const [extraFeeNote, setExtraFeeNote] = useState("");
  const [extraFeeCostCenter, setExtraFeeCostCenter] = useState("accounting");

  function resetAllocFields(c) {
    const a = (c && c.allocation) || {};
    setAllocAccounting(c ? String(a.accounting || "") : "");
    setAllocTax(c ? String(a.tax || "") : "");
    setAllocPayroll(c ? String(a.payroll || "") : "");
    setAllocOther(c ? String(a.other || "") : "");
    const sp = (c && c.feeSplit) || {};
    setFeeSplit(Object.fromEntries(COST_CENTERS.map((cc) => [cc, c && sp[cc] ? String(sp[cc]) : ""])));
    setFeeSplitNote(c ? (c.feeSplitNote || "") : "");
  }
  // Builds the €-split from the form and checks it against the fixed fee.
  // Returns { split, note } or null (after telling the user what's wrong).
  function validatedFeeSplit(fee) {
    const split = Object.fromEntries(COST_CENTERS.map((cc) => [cc, Number(feeSplit[cc]) || 0]));
    const note = feeSplitNote.trim();
    const total = feeSplitTotal(split);
    if (fee && total > 0 && Math.abs(total - fee) > 0.01) {
      window.alert(`The cost-center split adds up to ${fmtEur(total)}, but the fixed fee is ${fmtEur(fee)}. Adjust the split so it matches the fee (or leave all split fields empty to record the fee as unallocated for now).`);
      return null;
    }
    if (!fee && total > 0) {
      window.alert("You entered a cost-center split but no fixed fee. Enter the fixed fee first.");
      return null;
    }
    if (split.other > 0 && !note) {
      window.alert('Please add a note explaining what the "Other" part of the fee covers.');
      return null;
    }
    return { split, note: split.other > 0 ? note : "" };
  }
  const feeSplitRow = (split, note) => Object.assign(
    Object.fromEntries(COST_CENTERS.map((cc) => [`cc_${cc}`, split[cc] || 0])),
    { cc_other_note: note || null },
  );
  function openNew() {
    setModal("new"); setName(""); setFixedFee(""); setFeeEffectiveDate(toKey(TODAY)); setEndDate(""); setEditingHistoryId(null);
    resetAllocFields(null);
  }
  function openEdit(c) {
    setModal(c); setName(c.name); setFixedFee(c.fixedFee ?? ""); setFeeEffectiveDate(toKey(TODAY)); setEndDate(c.endDate || ""); setEditingHistoryId(null);
    resetAllocFields(c);
  }

  async function save() {
    if (!name.trim()) return;
    const fee = fixedFee === "" ? null : Number(fixedFee);
    const allocation = {
      accounting: Number(allocAccounting) || 0,
      tax: Number(allocTax) || 0,
      payroll: Number(allocPayroll) || 0,
      other: Number(allocOther) || 0,
    };
    const splitResult = validatedFeeSplit(fee);
    if (!splitResult) return;
    const { split, note: splitNote } = splitResult;
    if (modal === "new") {
      const { data, error } = await supabase.from("clients").insert({ name: name.trim(), fixed_fee: fee, end_date: endDate || null }).select().single();
      if (error || !data) {
        window.alert(`Couldn't create client: ${error ? error.message : "unknown error"}`);
        return;
      }
      const { error: feeError } = await supabase.from("client_fee_history").insert({
        client_id: data.id, effective_date: feeEffectiveDate || toKey(TODAY), fixed_fee: fee,
        alloc_accounting: allocation.accounting, alloc_tax: allocation.tax, alloc_payroll: allocation.payroll, alloc_other: allocation.other,
        ...feeSplitRow(split, splitNote),
      });
      if (feeError) {
        window.alert(`Client was created, but saving its fee terms failed: ${feeError.message}`);
      }
    } else {
      const { error } = await supabase.from("clients").update({ name: name.trim(), fixed_fee: fee, end_date: endDate || null }).eq("id", modal.id);
      if (error) {
        window.alert(`Couldn't update client: ${error.message}`);
        return;
      }
      const prevSplit = Object.fromEntries(COST_CENTERS.map((cc) => [cc, (modal.feeSplit && modal.feeSplit[cc]) || 0]));
      const prevFee = { fixedFee: modal.fixedFee ?? null, allocation: modal.allocation || { accounting: 0, tax: 0, payroll: 0, other: 0 }, feeSplit: prevSplit, feeSplitNote: modal.feeSplitNote || "" };
      const changed = JSON.stringify(prevFee) !== JSON.stringify({ fixedFee: fee, allocation, feeSplit: split, feeSplitNote: splitNote });
      if (changed) {
        const { error: feeError } = await supabase.from("client_fee_history").insert({
          client_id: modal.id, effective_date: feeEffectiveDate || toKey(TODAY), fixed_fee: fee,
          alloc_accounting: allocation.accounting, alloc_tax: allocation.tax, alloc_payroll: allocation.payroll, alloc_other: allocation.other,
          ...feeSplitRow(split, splitNote),
        });
        if (feeError) {
          window.alert(`Client was updated, but saving the new fee terms failed: ${feeError.message}`);
        }
      }
    }
    await refetchClients();
    setModal(null);
  }
  async function toggleActive(c) {
    await supabase.from("clients").update({ active: !c.active }).eq("id", c.id);
    refetchClients();
  }
  async function deleteClient(c) {
    if (!window.confirm(`Permanently delete "${c.name}"? This cannot be undone.`)) return;
    const { error } = await supabase.from("clients").delete().eq("id", c.id);
    if (error) {
      if (error.code === "23503") {
        window.alert(`Can't delete "${c.name}" — it still has logged hours on record. Use Deactivate instead, or remove its time entries first.`);
      } else {
        window.alert(`Couldn't delete "${c.name}": ${error.message}`);
      }
      return;
    }
    refetchClients();
  }

  function startEditHistory(h) {
    setEditingHistoryId(h.id);
    const a = h.allocation || {};
    const sp = h.feeSplit || {};
    setHistForm({
      effectiveDate: h.effectiveDate, fixedFee: h.fixedFee ?? "",
      accounting: a.accounting ?? "", tax: a.tax ?? "", payroll: a.payroll ?? "", other: a.other ?? "",
      split: Object.fromEntries(COST_CENTERS.map((cc) => [cc, sp[cc] ? String(sp[cc]) : ""])),
      splitNote: h.feeSplitNote || "",
    });
  }
  function cancelEditHistory() { setEditingHistoryId(null); setHistForm(null); }
  async function saveHistoryEdit() {
    const fee = histForm.fixedFee === "" ? null : Number(histForm.fixedFee);
    const split = Object.fromEntries(COST_CENTERS.map((cc) => [cc, Number(histForm.split?.[cc]) || 0]));
    const total = feeSplitTotal(split);
    if (fee && total > 0 && Math.abs(total - fee) > 0.01) {
      window.alert(`The cost-center split adds up to ${fmtEur(total)}, but the fixed fee is ${fmtEur(fee)}. Adjust the split so it matches.`);
      return;
    }
    if (split.other > 0 && !(histForm.splitNote || "").trim()) {
      window.alert('Please add a note explaining what the "Other" part of the fee covers.');
      return;
    }
    await supabase.from("client_fee_history").update({
      effective_date: histForm.effectiveDate,
      fixed_fee: fee,
      alloc_accounting: Number(histForm.accounting) || 0,
      alloc_tax: Number(histForm.tax) || 0,
      alloc_payroll: Number(histForm.payroll) || 0,
      alloc_other: Number(histForm.other) || 0,
      ...feeSplitRow(split, split.other > 0 ? histForm.splitNote.trim() : ""),
    }).eq("id", editingHistoryId);
    await refetchClients();
    cancelEditHistory();
  }
  async function deleteHistoryEntry(entryId) {
    await supabase.from("client_fee_history").delete().eq("id", entryId);
    await refetchClients();
    if (editingHistoryId === entryId) cancelEditHistory();
  }

  function openExtraFee(clientId, defaultMonth) {
    setExtraFeeClientId(clientId);
    setExtraFeeEditingId(null);
    setExtraFeeMonth(defaultMonth);
    setExtraFeeAmount("");
    setExtraFeeNote("");
    setExtraFeeCostCenter("accounting");
  }
  // Edit an existing extra fee (month, amount, cost center, note) — same
  // modal as "add"; saving updates the row in place.
  function openEditExtraFee(clientId, x) {
    setExtraFeeClientId(clientId);
    setExtraFeeEditingId(x.id);
    setExtraFeeMonth(x.month);
    setExtraFeeAmount(String(x.amount));
    setExtraFeeNote(x.note || "");
    setExtraFeeCostCenter(COST_CENTERS.includes(x.costCenter) ? x.costCenter : "other");
  }
  async function saveExtraFee() {
    if (!extraFeeMonth || !extraFeeAmount || Number(extraFeeAmount) <= 0) return;
    if (extraFeeCostCenter === "other" && !extraFeeNote.trim()) {
      window.alert('A note is required when the cost center is "Other" — say what this fee is for.');
      return;
    }
    const row = { month: extraFeeMonth, amount: Number(extraFeeAmount), note: extraFeeNote.trim() || null, cost_center: extraFeeCostCenter };
    const { error } = extraFeeEditingId
      ? await supabase.from("client_extra_fees").update(row).eq("id", extraFeeEditingId)
      : await supabase.from("client_extra_fees").insert({ client_id: extraFeeClientId, ...row });
    if (error) { window.alert(`Couldn't save extra fee: ${error.message}`); return; }
    await refetchClients();
    setExtraFeeClientId(null);
    setExtraFeeEditingId(null);
  }
  async function deleteExtraFee(entryId) {
    await supabase.from("client_extra_fees").delete().eq("id", entryId);
    refetchClients();
  }

  const [viewRangeStart, viewRangeEnd] = periodRange(viewPeriod, viewAnchor);
  const periodEntries = entries.filter((e) => { const d = fromKey(e.date); return d >= viewRangeStart && d <= viewRangeEnd; });
  const filtered = clients.filter((c) => c.name.toLowerCase().includes(query.toLowerCase()));
  const liveClient = modal && modal !== "new" ? clients.find((c) => c.id === modal.id) : null;

  return (
    <div>
      <PageHeader title="Clients" sub="Manage the clients your team logs hours against, and their billing terms."
        right={
          <div style={{ display: "flex", gap: 8 }}>
            <Btn variant="secondary" onClick={() => exportToExcel(`clients-${toKey(TODAY)}.xlsx`, [{
              name: "Clients",
              rows: clients.map((c) => ({
                Name: c.name, Status: c.active ? "Active" : "Inactive",
                "Fixed fee (€/mo)": c.fixedFee ?? "", "Accounting alloc. (h)": c.allocation?.accounting || 0,
                "Tax alloc. (h)": c.allocation?.tax || 0, "Payroll alloc. (h)": c.allocation?.payroll || 0,
                "Other alloc. (h)": c.allocation?.other || 0,
                ...Object.fromEntries(COST_CENTERS.map((cc) => [`Fee → ${COST_CENTER_LABELS[cc]} (€)`, c.feeSplit?.[cc] || 0])),
                "Fee → Other note": c.feeSplitNote || "",
              })),
            }])}>Export to Excel</Btn>
            <Btn icon={Plus} onClick={openNew}>Add client</Btn>
          </div>
        } />
      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 10 }}>
          Hours, allocation, and fee below reflect this reporting period — not just the latest settings.
        </div>
        <PeriodSelector period={viewPeriod} setPeriod={setViewPeriod} anchor={viewAnchor} setAnchor={setViewAnchor} />
      </Panel>
      <Panel style={{ marginBottom: 18 }}>
        <div style={{ marginBottom: 14, position: "relative", maxWidth: 260 }}>
          <Search size={14} color={C.inkFaint} style={{ position: "absolute", left: 9, top: 9 }} />
          <input placeholder="Search clients…" value={query} onChange={(e) => setQuery(e.target.value)}
            style={{ ...inputStyle, width: "100%", paddingLeft: 28 }} />
        </div>
        <TableShell headers={["Client", "Hours", "Revenue", "Allocation", "Fixed fee", "Extra fees (period)", "Status", ""]}>
          {filtered.map((c) => {
            const clientPeriodEntries = periodEntries.filter((e) => e.clientId === c.id);
            const hrs = clientPeriodEntries.reduce((s, e) => s + e.hours, 0);
            const { revenue, allocation } = periodClientMetrics(c, clientPeriodEntries, viewRangeStart, viewRangeEnd);
            const feeThen = feeAsOf(c, toKey(viewRangeStart));
            const a = feeThen.allocation || {};
            const allocTitle = CATEGORIES.map((cat) => `${CATEGORY_LABELS[cat]} ${a[cat] || 0}h`).join(" · ");
            const extraThisPeriod = extraFeesAccruedForRange(c, viewRangeStart, viewRangeEnd);
            const defaultMonth = `${viewRangeStart.getFullYear()}-${pad(viewRangeStart.getMonth() + 1)}`;
            return (
              <tr key={c.id} style={{ opacity: c.active ? 1 : 0.55 }}>
                <Td>{c.name}</Td>
                <Td mono>{hrs.toFixed(2)}h</Td>
                <Td mono>{fmtEur(revenue)}</Td>
                <Td mono title={allocTitle}>{allocation ? allocation.toFixed(1) + "h" : "—"}</Td>
                <Td mono title={feeThen.fixedFee ? (feeSplitTotal(feeThen.feeSplit) > 0 ? COST_CENTERS.filter((cc) => feeThen.feeSplit[cc]).map((cc) => `${COST_CENTER_LABELS[cc]} ${fmtEur(feeThen.feeSplit[cc])}`).join(" · ") : "No cost-center split yet") : undefined}>
                  {feeThen.fixedFee ? fmtEur(feeThen.fixedFee) + "/mo" : "—"}
                  {feeThen.fixedFee && feeSplitTotal(feeThen.feeSplit) <= 0 ? <span title="Fixed fee has no cost-center split yet" style={{ marginLeft: 6, color: C.warn }}>●</span> : null}
                </Td>
                <Td mono>
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span>{extraThisPeriod > 0 ? fmtEur(extraThisPeriod) : "—"}</span>
                    <button onClick={() => openExtraFee(c.id, defaultMonth)} title="Add extra fee"
                      style={{ ...iconBtnStyle, border: `1px solid ${C.border}`, borderRadius: 5, padding: "2px 5px" }}>
                      <Plus size={11} color={C.accent} />
                    </button>
                  </div>
                </Td>
                <Td>
                  {c.endDate && c.endDate <= toKey(TODAY)
                    ? <Badge tone="warn">Ended {c.endDate}</Badge>
                    : c.active ? <Badge tone="accent">Active</Badge> : <Badge tone="neutral">Inactive</Badge>}
                </Td>
                <Td right>
                  <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                    <Btn size="sm" variant="ghost" icon={Pencil} onClick={() => openEdit(c)}>Edit</Btn>
                    <Btn size="sm" variant={c.active ? "danger" : "secondary"} onClick={() => toggleActive(c)}>
                      {c.active ? "Deactivate" : "Reactivate"}
                    </Btn>
                    <Btn size="sm" variant="ghost" icon={Trash2} onClick={() => deleteClient(c)}>Delete</Btn>
                  </div>
                </Td>
              </tr>
            );
          })}
          {(() => {
            // Totals across every client currently listed (respects the
            // search box), for this reporting period.
            let totalHrs = 0, totalRevenue = 0, totalExtra = 0, totalFixedFee = 0;
            filtered.forEach((c) => {
              const clientPeriodEntries = periodEntries.filter((e) => e.clientId === c.id);
              totalHrs += clientPeriodEntries.reduce((s, e) => s + e.hours, 0);
              totalRevenue += periodClientMetrics(c, clientPeriodEntries, viewRangeStart, viewRangeEnd).revenue;
              totalExtra += extraFeesAccruedForRange(c, viewRangeStart, viewRangeEnd);
              const feeThen = feeAsOf(c, toKey(viewRangeStart));
              if (feeThen.fixedFee && (!c.endDate || c.endDate > toKey(viewRangeStart))) totalFixedFee += feeThen.fixedFee;
            });
            return (
              <tr style={{ fontWeight: 700, borderTop: `2px solid ${C.borderStrong}` }}>
                <Td>Total ({filtered.length} client{filtered.length === 1 ? "" : "s"})</Td>
                <Td mono>{totalHrs.toFixed(2)}h</Td>
                <Td mono>{fmtEur(totalRevenue)}</Td>
                <Td />
                <Td mono>{fmtEur(totalFixedFee)}/mo</Td>
                <Td mono>{fmtEur(totalExtra)}</Td>
                <Td /><Td />
              </tr>
            );
          })()}
        </TableShell>
      </Panel>

      <Panel>
        <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 4 }}>Category workload</div>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Allocated vs. actual hours per category, for this period.</div>
        <TableShell headers={["Client", ...CATEGORIES.flatMap((c) => [`${CATEGORY_LABELS[c]} alloc.`, `${CATEGORY_LABELS[c]} actual`])]}>
          {filtered.filter((c) => c.active).map((c) => {
            const clientPeriodEntries = periodEntries.filter((e) => e.clientId === c.id);
            const { allocationByCategory, actualByCategory } = periodClientMetrics(c, clientPeriodEntries, viewRangeStart, viewRangeEnd);
            return (
              <tr key={c.id}>
                <Td>{c.name}</Td>
                {CATEGORIES.map((cat) => (
                  <React.Fragment key={cat}>
                    <Td mono>{allocationByCategory[cat] > 0 ? allocationByCategory[cat].toFixed(1) + "h" : "—"}</Td>
                    <Td mono style={{ color: actualByCategory[cat] > allocationByCategory[cat] && allocationByCategory[cat] > 0 ? C.warn : undefined }}>
                      {actualByCategory[cat] > 0 ? actualByCategory[cat].toFixed(1) + "h" : "—"}
                    </Td>
                  </React.Fragment>
                ))}
              </tr>
            );
          })}
        </TableShell>
      </Panel>

      {modal && (
        <Modal title={modal === "new" ? "Add client" : "Edit client"} onClose={() => setModal(null)} width={560}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Field label="Client name">
              <input style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Acme Retail" />
            </Field>
            {liveClient && (() => {
              const clientEntries = entries.filter((e) => e.clientId === liveClient.id);
              const [pStart, pEnd] = periodRange(viewPeriod, viewAnchor);
              const inRange = (rStart, rEnd) => clientEntries.filter((e) => { const d = fromKey(e.date); return d >= rStart && d <= rEnd; });
              const periodEntriesForClient = inRange(pStart, pEnd);
              const periodRevenue = periodClientMetrics(liveClient, periodEntriesForClient, pStart, pEnd).revenue;
              const periodHours = periodEntriesForClient.reduce((s, e) => s + e.hours, 0);
              const yearStart = new Date(TODAY.getFullYear(), 0, 1);
              const ytdEntries = inRange(yearStart, TODAY);
              const ytdRevenue = periodClientMetrics(liveClient, ytdEntries, yearStart, TODAY).revenue;
              const ytdHours = ytdEntries.reduce((s, e) => s + e.hours, 0);
              const startKey = earliestHistoryDate(liveClient.feeHistory);
              const firstEntryDate = clientEntries.reduce((min, e) => { const d = fromKey(e.date); return !min || d < min ? d : min; }, null);
              const allStart = startKey ? fromKey(startKey) : (firstEntryDate || TODAY);
              const allEndRaw = liveClient.endDate ? fromKey(liveClient.endDate) : TODAY;
              const allEnd = allEndRaw < TODAY ? allEndRaw : TODAY;
              const allEntries = inRange(allStart, allEnd);
              const allRevenue = periodClientMetrics(liveClient, allEntries, allStart, allEnd).revenue;
              const allHours = clientEntries.reduce((s, e) => s + e.hours, 0);
              const rows = [
                { label: periodLabel(viewPeriod, pStart, pEnd), hours: periodHours, revenue: periodRevenue },
                { label: `Year to date (${TODAY.getFullYear()})`, hours: ytdHours, revenue: ytdRevenue },
                { label: "All time", hours: allHours, revenue: allRevenue },
              ];
              return (
                <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, padding: "10px 12px" }}>
                  <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink, marginBottom: 8 }}>Summary</div>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: sans, fontSize: 12.5 }}>
                    <thead>
                      <tr style={{ color: C.inkFaint, textAlign: "left" }}>
                        <th style={{ fontWeight: 500, paddingBottom: 4 }}></th>
                        <th style={{ fontWeight: 500, paddingBottom: 4, textAlign: "right" }}>Hours</th>
                        <th style={{ fontWeight: 500, paddingBottom: 4, textAlign: "right" }}>Revenue</th>
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((r) => (
                        <tr key={r.label}>
                          <td style={{ padding: "2px 0", color: C.inkMuted }}>{r.label}</td>
                          <td style={{ padding: "2px 0", fontFamily: mono, color: C.ink, textAlign: "right" }}>{r.hours.toFixed(2)}h</td>
                          <td style={{ padding: "2px 0", fontFamily: mono, color: C.ink, textAlign: "right" }}>{fmtEur(r.revenue)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              );
            })()}
            <div>
              <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink, marginBottom: 8 }}>Agreed monthly allocation, by category</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <Field label="Accounting hours">
                  <input type="number" min="0" style={inputStyle} value={allocAccounting} onChange={(e) => setAllocAccounting(e.target.value)} placeholder="e.g. 20" />
                </Field>
                <Field label="Tax hours">
                  <input type="number" min="0" style={inputStyle} value={allocTax} onChange={(e) => setAllocTax(e.target.value)} placeholder="e.g. 15" />
                </Field>
                <Field label="Payroll hours">
                  <input type="number" min="0" style={inputStyle} value={allocPayroll} onChange={(e) => setAllocPayroll(e.target.value)} placeholder="e.g. 5" />
                </Field>
                <Field label="Other hours">
                  <input type="number" min="0" style={inputStyle} value={allocOther} onChange={(e) => setAllocOther(e.target.value)} placeholder="e.g. 5" />
                </Field>
              </div>
            </div>
            <Field label="Fixed monthly fee — retainer (optional)">
              <input type="number" min="0" style={inputStyle} value={fixedFee}
                onChange={(e) => setFixedFee(e.target.value)} placeholder="e.g. 1200" />
            </Field>
            {(() => {
              const fee = fixedFee === "" ? 0 : Number(fixedFee) || 0;
              const total = feeSplitTotal(feeSplit);
              const remaining = fee - total;
              const tone = total === 0 ? C.inkFaint : Math.abs(remaining) <= 0.01 ? C.accent : C.danger;
              return (
                <div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 8 }}>
                    <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink }}>Fixed fee split by cost center (€ per month)</div>
                    <div style={{ fontFamily: mono, fontSize: 11.5, color: tone }}>
                      {total === 0 ? "No split yet" : Math.abs(remaining) <= 0.01 ? "Matches fee ✓" : remaining > 0 ? `${fmtEur(remaining)} left to allocate` : `${fmtEur(-remaining)} over the fee`}
                    </div>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10 }}>
                    {COST_CENTERS.map((cc) => (
                      <Field key={cc} label={COST_CENTER_LABELS[cc]}>
                        <input type="number" min="0" style={inputStyle} value={feeSplit[cc]}
                          onChange={(e) => setFeeSplit((f) => ({ ...f, [cc]: e.target.value }))} placeholder="0" />
                      </Field>
                    ))}
                  </div>
                  {Number(feeSplit.other) > 0 && (
                    <div style={{ marginTop: 10 }}>
                      <Field label='Note — what does the "Other" part cover? (required)'>
                        <input style={inputStyle} value={feeSplitNote} onChange={(e) => setFeeSplitNote(e.target.value)} placeholder="e.g. Company secretarial work" />
                      </Field>
                    </div>
                  )}
                </div>
              );
            })()}
            <Field label="Effective from">
              <input type="date" style={inputStyle} value={feeEffectiveDate} onChange={(e) => setFeeEffectiveDate(e.target.value)} />
            </Field>
            <Field label="End date — client relationship ended (optional)">
              <input type="date" style={inputStyle} value={endDate} onChange={(e) => setEndDate(e.target.value)} />
            </Field>
            <div style={{ fontFamily: sans, fontSize: 11.5, color: C.inkFaint, lineHeight: 1.5 }}>
              Revenue is the fixed fee (this recurring rate) plus any extra fees added below for
              specific months. The cost-center split says how the fixed fee is reported on the
              dashboard (the split must add up to the fee); the hours allocation is for tracking
              workload only (allocated vs. actual). Saving with different numbers above adds a new
              row below, effective from that date — it never overwrites past periods.
            </div>
            <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink }}>Extra fees by month</div>
                {liveClient && (
                  <Btn size="sm" variant="secondary" icon={Plus} onClick={() => openExtraFee(liveClient.id, toKey(TODAY).slice(0, 7))}>Add</Btn>
                )}
              </div>
              {liveClient && liveClient.extraFeeEntries && liveClient.extraFeeEntries.length > 0 ? (
                <div className="scrollbar-thin" style={{ maxHeight: 140, overflowY: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: sans, fontSize: 12 }}>
                    <thead>
                      <tr>
                        {["Month", "Amount", "Cost center", "Note", ""].map((h) => (
                          <th key={h} style={{ textAlign: "left", padding: "3px 6px", fontSize: 10.5, fontWeight: 600, color: C.inkFaint, borderBottom: `1px solid ${C.border}` }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {[...liveClient.extraFeeEntries].sort((a, b) => b.month.localeCompare(a.month)).map((x) => (
                        <tr key={x.id}>
                          <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{x.month}</td>
                          <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{fmtEur(x.amount)}</td>
                          <td style={{ padding: "3px 6px", color: C.ink }}>{costCenterLabel(x.costCenter)}</td>
                          <td style={{ padding: "3px 6px", color: C.ink }}>{x.note || "—"}</td>
                          <td style={{ padding: "3px 6px" }}>
                            <div style={{ display: "flex", gap: 4 }}>
                              <button onClick={() => openEditExtraFee(liveClient.id, x)} style={iconBtnStyle} title="Edit"><Pencil size={12} color={C.inkMuted} /></button>
                              <button onClick={() => deleteExtraFee(x.id)} style={iconBtnStyle} title="Delete"><Trash2 size={12} color={C.inkMuted} /></button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div style={{ fontFamily: sans, fontSize: 12, color: C.inkFaint }}>No extra fees added yet.</div>
              )}
            </div>
            {liveClient && liveClient.feeHistory && liveClient.feeHistory.length > 0 && (
              <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 12 }}>
                <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink, marginBottom: 8 }}>Fee history</div>
                <div className="scrollbar-thin" style={{ maxHeight: 220, overflowY: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: sans, fontSize: 12 }}>
                    <thead>
                      <tr>
                        {["From", "To", "Alloc. total", "Fixed fee", "Split", ""].map((h) => (
                          <th key={h} style={{ textAlign: "left", padding: "3px 6px", fontSize: 10.5, fontWeight: 600, color: C.inkFaint, borderBottom: `1px solid ${C.border}` }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {[...withPeriods(liveClient.feeHistory)].reverse().map((h) => {
                        const a = h.allocation || {};
                        const allocTotal = CATEGORIES.reduce((s, c) => s + (a[c] || 0), 0);
                        const allocTitle = CATEGORIES.map((cat) => `${CATEGORY_LABELS[cat]} ${a[cat] || 0}h`).join(" · ");
                        const sp = h.feeSplit || {};
                        const spTotal = feeSplitTotal(sp);
                        const spTitle = COST_CENTERS.filter((cc) => sp[cc]).map((cc) => `${COST_CENTER_LABELS[cc]} ${fmtEur(sp[cc])}`).join(" · ") + (h.feeSplitNote ? ` — ${h.feeSplitNote}` : "");
                        const isEditing = editingHistoryId === h.id;
                        return isEditing ? (
                          <tr key={h.id}>
                            <td colSpan={6} style={{ padding: "6px" }}>
                              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                                <input type="date" style={{ ...inputStyle, width: 130, padding: "4px 6px" }} value={histForm.effectiveDate} onChange={(ev) => setHistForm((f) => ({ ...f, effectiveDate: ev.target.value }))} />
                                <input type="number" style={{ ...inputStyle, width: 65, padding: "4px 6px" }} value={histForm.accounting} onChange={(ev) => setHistForm((f) => ({ ...f, accounting: ev.target.value }))} placeholder="Acct h" />
                                <input type="number" style={{ ...inputStyle, width: 65, padding: "4px 6px" }} value={histForm.tax} onChange={(ev) => setHistForm((f) => ({ ...f, tax: ev.target.value }))} placeholder="Tax h" />
                                <input type="number" style={{ ...inputStyle, width: 65, padding: "4px 6px" }} value={histForm.payroll} onChange={(ev) => setHistForm((f) => ({ ...f, payroll: ev.target.value }))} placeholder="Payroll h" />
                                <input type="number" style={{ ...inputStyle, width: 65, padding: "4px 6px" }} value={histForm.other} onChange={(ev) => setHistForm((f) => ({ ...f, other: ev.target.value }))} placeholder="Other h" />
                                <input type="number" style={{ ...inputStyle, width: 80, padding: "4px 6px" }} value={histForm.fixedFee} onChange={(ev) => setHistForm((f) => ({ ...f, fixedFee: ev.target.value }))} placeholder="Fixed fee" />
                              </div>
                              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center", marginTop: 6 }}>
                                <span style={{ fontFamily: sans, fontSize: 11, color: C.inkFaint, width: 130 }}>Fee split (€):</span>
                                {COST_CENTERS.map((cc) => (
                                  <input key={cc} type="number" style={{ ...inputStyle, width: 65, padding: "4px 6px" }} value={histForm.split?.[cc] ?? ""}
                                    onChange={(ev) => setHistForm((f) => ({ ...f, split: { ...(f.split || {}), [cc]: ev.target.value } }))} placeholder={COST_CENTER_LABELS[cc]} title={COST_CENTER_LABELS[cc]} />
                                ))}
                                {Number(histForm.split?.other) > 0 && (
                                  <input style={{ ...inputStyle, width: 160, padding: "4px 6px" }} value={histForm.splitNote || ""} onChange={(ev) => setHistForm((f) => ({ ...f, splitNote: ev.target.value }))} placeholder="Other — note (required)" />
                                )}
                                <button onClick={() => saveHistoryEdit()} style={iconBtnStyle}><Check size={14} color={C.accent} /></button>
                                <button onClick={cancelEditHistory} style={iconBtnStyle}><X size={14} color={C.inkMuted} /></button>
                              </div>
                            </td>
                          </tr>
                        ) : (
                          <tr key={h.id}>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.from}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.to}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }} title={allocTitle}>{allocTotal > 0 ? allocTotal + "h" : "—"}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.fixedFee ? fmtEur(h.fixedFee) : "—"}</td>
                            <td style={{ padding: "3px 6px", fontFamily: sans, color: spTotal > 0 ? C.ink : C.warn, fontSize: 11 }} title={spTitle || undefined}>
                              {spTotal > 0 ? COST_CENTERS.filter((cc) => sp[cc]).map((cc) => COST_CENTER_LABELS[cc].slice(0, 3)).join("/") : (h.fixedFee ? "Not split" : "—")}
                            </td>
                            <td style={{ padding: "3px 6px" }}>
                              <div style={{ display: "flex", gap: 4 }}>
                                <button onClick={() => startEditHistory(h)} style={iconBtnStyle}><Pencil size={12} color={C.inkMuted} /></button>
                                <button onClick={() => deleteHistoryEntry(h.id)} style={iconBtnStyle}><Trash2 size={12} color={C.inkMuted} /></button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 6 }}>
              <Btn variant="secondary" onClick={() => setModal(null)}>Cancel</Btn>
              <Btn onClick={save}>Save client</Btn>
            </div>
          </div>
        </Modal>
      )}

      {extraFeeClientId && (
        <Modal title={`${extraFeeEditingId ? "Edit" : "Add"} extra fee — ${clients.find((c) => c.id === extraFeeClientId)?.name || ""}`} onClose={() => { setExtraFeeClientId(null); setExtraFeeEditingId(null); }} width={360}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Field label="Month">
              <input type="month" style={inputStyle} value={extraFeeMonth} onChange={(e) => setExtraFeeMonth(e.target.value)} />
            </Field>
            <Field label="Amount">
              <input type="number" min="0" style={inputStyle} value={extraFeeAmount} onChange={(e) => setExtraFeeAmount(e.target.value)} placeholder="e.g. 150" autoFocus />
            </Field>
            <Field label="Cost center">
              <select style={inputStyle} value={extraFeeCostCenter} onChange={(e) => setExtraFeeCostCenter(e.target.value)}>
                {COST_CENTERS.map((cc) => <option key={cc} value={cc}>{COST_CENTER_LABELS[cc]}</option>)}
              </select>
            </Field>
            <Field label={extraFeeCostCenter === "other" ? "Note (required for Other)" : "Note (optional)"}>
              <input style={inputStyle} value={extraFeeNote} onChange={(e) => setExtraFeeNote(e.target.value)} placeholder="e.g. Extra VAT filing" />
            </Field>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 6 }}>
              <Btn variant="secondary" onClick={() => setExtraFeeClientId(null)}>Cancel</Btn>
              <Btn onClick={saveExtraFee}>{extraFeeEditingId ? "Save changes" : "Add fee"}</Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}


/* ---------------------------------------------------------------------- */
/* Admin: Employees                                                        */
/* ---------------------------------------------------------------------- */

function AdminEmployees({ employees, refetchEmployees, entries }) {
  const [modal, setModal] = useState(null); // null | employee object (accounts are created via self sign-up)
  const [name, setName] = useState("");
  const [title, setTitle] = useState("");
  const [role, setRole] = useState("employee");
  const [weeklyHours, setWeeklyHours] = useState("40");
  const [annualLeaveDays, setAnnualLeaveDays] = useState("25");
  const [department, setDepartment] = useState(DEFAULT_DEPARTMENT);
  const [endDate, setEndDate] = useState("");
  const [grossSalary, setGrossSalary] = useState("");
  const [socialSecurity, setSocialSecurity] = useState("");
  const [ticketRestaurant, setTicketRestaurant] = useState("");
  const [insurance, setInsurance] = useState("");
  const [otherCost, setOtherCost] = useState("");
  const [costEffectiveDate, setCostEffectiveDate] = useState(toKey(TODAY));
  const [editingHistoryId, setEditingHistoryId] = useState(null);
  const [histForm, setHistForm] = useState(null);
  const [viewPeriod, setViewPeriod] = useState("month");
  const [viewAnchor, setViewAnchor] = useState(TODAY);

  function resetCostFields(e) {
    setGrossSalary(e ? String(e.grossSalary || "") : "");
    setSocialSecurity(e ? String(e.socialSecurity || "") : "");
    setTicketRestaurant(e ? String(e.ticketRestaurant || "") : "");
    setInsurance(e ? String(e.insurance || "") : "");
    setOtherCost(e ? String(e.otherCost || "") : "");
    setCostEffectiveDate(toKey(TODAY));
    setEditingHistoryId(null);
  }
  function openEdit(e) {
    setModal(e); setName(e.name); setTitle(e.title || ""); setRole(e.role);
    setWeeklyHours(String(e.weeklyHours || 40)); setAnnualLeaveDays(String(e.annualLeaveDays || 25));
    setDepartment(e.department || DEFAULT_DEPARTMENT);
    setEndDate(e.endDate || "");
    resetCostFields(e);
  }

  async function save() {
    if (!modal || !name.trim()) return;
    const hrs = Number(weeklyHours) || 40;
    const leaveDays = Number(annualLeaveDays) || 25;
    const costFields = {
      grossSalary: Number(grossSalary) || 0,
      socialSecurity: Number(socialSecurity) || 0,
      ticketRestaurant: Number(ticketRestaurant) || 0,
      insurance: Number(insurance) || 0,
      otherCost: Number(otherCost) || 0,
    };
    const { error: profileError } = await supabase.from("profiles").update({
      name: name.trim(), title: title.trim() || null, role,
      weekly_hours: hrs, annual_leave_days: leaveDays,
      department, end_date: endDate || null,
    }).eq("id", modal.id);
    if (profileError) {
      window.alert(`Couldn't update employee: ${profileError.message}`);
      return;
    }

    const prevCost = { grossSalary: modal.grossSalary || 0, socialSecurity: modal.socialSecurity || 0, ticketRestaurant: modal.ticketRestaurant || 0, insurance: modal.insurance || 0, otherCost: modal.otherCost || 0 };
    const changed = JSON.stringify(prevCost) !== JSON.stringify(costFields);
    if (changed) {
      await supabase.from("employee_cost_history").insert({
        employee_id: modal.id, effective_date: costEffectiveDate || toKey(TODAY),
        gross_salary: costFields.grossSalary, social_security: costFields.socialSecurity,
        ticket_restaurant: costFields.ticketRestaurant, insurance: costFields.insurance, other_cost: costFields.otherCost,
      });
    }
    await refetchEmployees();
    setModal(null);
  }
  async function toggleActive(e) {
    await supabase.from("profiles").update({ active: !e.active }).eq("id", e.id);
    refetchEmployees();
  }
  async function deleteEmployee(e) {
    if (!window.confirm(`Permanently delete "${e.name}"? This removes their HoursLedger record but does not revoke their login — this cannot be undone.`)) return;
    const { error } = await supabase.from("profiles").delete().eq("id", e.id);
    if (error) {
      if (error.code === "23503") {
        window.alert(`Can't delete "${e.name}" — they still have logged hours or leave requests on record. Use Deactivate instead.`);
      } else {
        window.alert(`Couldn't delete "${e.name}": ${error.message}`);
      }
      return;
    }
    refetchEmployees();
  }

  function startEditHistory(h) {
    setEditingHistoryId(h.id);
    setHistForm({ effectiveDate: h.effectiveDate, grossSalary: String(h.grossSalary || ""), socialSecurity: String(h.socialSecurity || ""), ticketRestaurant: String(h.ticketRestaurant || ""), insurance: String(h.insurance || ""), otherCost: String(h.otherCost || "") });
  }
  function cancelEditHistory() { setEditingHistoryId(null); setHistForm(null); }
  async function saveHistoryEdit() {
    await supabase.from("employee_cost_history").update({
      effective_date: histForm.effectiveDate,
      gross_salary: Number(histForm.grossSalary) || 0,
      social_security: Number(histForm.socialSecurity) || 0,
      ticket_restaurant: Number(histForm.ticketRestaurant) || 0,
      insurance: Number(histForm.insurance) || 0,
      other_cost: Number(histForm.otherCost) || 0,
    }).eq("id", editingHistoryId);
    await refetchEmployees();
    cancelEditHistory();
  }
  async function deleteHistoryEntry(entryId) {
    await supabase.from("employee_cost_history").delete().eq("id", entryId);
    await refetchEmployees();
    if (editingHistoryId === entryId) cancelEditHistory();
  }

  const [viewRangeStart, viewRangeEnd] = periodRange(viewPeriod, viewAnchor);
  const periodEntries = entries.filter((e) => { const d = fromKey(e.date); return d >= viewRangeStart && d <= viewRangeEnd; });
  const liveEmployee = modal ? employees.find((e) => e.id === modal.id) : null;

  return (
    <div>
      <PageHeader title="Employees" sub="People join by creating their own account on the sign-in screen. Set their role, contracted hours, and labor cost here."
        right={
          <Btn variant="secondary" onClick={() => exportToExcel(`employees-${toKey(TODAY)}.xlsx`, [{
            name: "Employees",
            rows: employees.map((e) => ({
              Name: e.name, Title: e.title || "", Email: e.email, Role: e.role, Status: e.active ? "Active" : "Inactive",
              "Contracted h/wk": e.weeklyHours || 40, "Annual leave (d/yr)": e.annualLeaveDays || 25,
              "Gross salary": e.grossSalary || 0, "Social security": e.socialSecurity || 0,
              "Ticket restaurant": e.ticketRestaurant || 0, "Other insurance": e.insurance || 0, "Other cost": e.otherCost || 0,
              Department: departmentLabel(e.department || DEFAULT_DEPARTMENT), "Departure date": e.endDate || "",
            })),
          }])}>Export to Excel</Btn>
        } />
      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 10 }}>
          Hours and cost below reflect this reporting period — cost uses whatever was in effect at the time, not just the latest figure.
        </div>
        <PeriodSelector period={viewPeriod} setPeriod={setViewPeriod} anchor={viewAnchor} setAnchor={setViewAnchor} />
      </Panel>
      <Panel>
        <TableShell headers={["Name", "Title", "Email", "Role", "Department", "Departed", "Contracted /wk", "Leave/yr", "Hours", "Cost /mo", "Cost /h", "Status", ""]}>
          {employees.map((e) => {
            const hrs = periodEntries.filter((t) => t.employeeId === e.id).reduce((s, t) => s + t.hours, 0);
            const costThenMonthly = monthlyLaborCostAsOf(e, toKey(viewRangeStart));
            const costThenHourly = hourlyLaborCostAsOf(e, toKey(viewRangeStart));
            return (
              <tr key={e.id} style={{ opacity: e.active ? 1 : 0.55 }}>
                <Td>{e.name}</Td>
                <Td>{e.title || "—"}</Td>
                <Td>{e.email}</Td>
                <Td><Badge tone={e.role === "admin" ? "accent" : "neutral"}>{e.role === "admin" ? "Manager" : "Employee"}</Badge></Td>
                <Td>{departmentLabel(e.department || DEFAULT_DEPARTMENT)}</Td>
                <Td mono style={{ color: e.endDate ? C.warn : undefined }}>{e.endDate || "—"}</Td>
                <Td mono>{e.weeklyHours || 40}h</Td>
                <Td mono>{e.annualLeaveDays || 25}d</Td>
                <Td mono>{hrs.toFixed(2) + "h"}</Td>
                <Td mono>{fmtEur(costThenMonthly)}</Td>
                <Td mono>{fmtEur(costThenHourly)}</Td>
                <Td>{e.active ? <Badge tone="accent">Active</Badge> : <Badge tone="neutral">Inactive</Badge>}</Td>
                <Td right>
                  <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                    <Btn size="sm" variant="ghost" icon={Pencil} onClick={() => openEdit(e)}>Edit</Btn>
                    <Btn size="sm" variant={e.active ? "danger" : "secondary"} onClick={() => toggleActive(e)}>
                      {e.active ? "Deactivate" : "Reactivate"}
                    </Btn>
                    <Btn size="sm" variant="ghost" icon={Trash2} onClick={() => deleteEmployee(e)}>Delete</Btn>
                  </div>
                </Td>
              </tr>
            );
          })}
        </TableShell>
      </Panel>

      {modal && (
        <Modal title="Edit employee" onClose={() => setModal(null)} width={520}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Field label="Full name">
              <input style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Maria Ioannou" />
            </Field>
            <Field label="Job title">
              <input style={inputStyle} value={title} onChange={(e) => setTitle(e.target.value)} placeholder="e.g. Senior Accountant" />
            </Field>
            <Field label="Contracted hours per week">
              <input type="number" min="1" max="80" style={inputStyle} value={weeklyHours} onChange={(e) => setWeeklyHours(e.target.value)} placeholder="e.g. 40" />
            </Field>
            <Field label="Annual leave entitlement (days/year)">
              <input type="number" min="0" max="40" style={inputStyle} value={annualLeaveDays} onChange={(e) => setAnnualLeaveDays(e.target.value)} placeholder="e.g. 25" />
            </Field>
            <Field label="Role">
              <select style={inputStyle} value={role} onChange={(e) => setRole(e.target.value)}>
                <option value="employee">Employee</option>
                <option value="admin">Manager / Admin</option>
              </select>
            </Field>
            <Field label="Department — where this person's labor cost is reported">
              <select style={inputStyle} value={department} onChange={(e) => setDepartment(e.target.value)}>
                {DEPARTMENTS.map((d) => <option key={d.key} value={d.key}>{d.label} ({d.costCenters.map((cc) => COST_CENTER_LABELS[cc]).join(" / ")})</option>)}
              </select>
            </Field>
            <div style={{ fontFamily: sans, fontSize: 11.5, color: C.inkFaint, lineHeight: 1.5, marginTop: -4 }}>
              Accounting Team cost is spread across Accounting / Tax / Payroll / Other by the hours logged each period. Management Team cost is split between Director and Finance in proportion to that period's Director/Finance fees. Rental goes entirely to the Rental cost center.
            </div>
            <Field label="Departure date (optional) — leave empty while still with the firm">
              <input type="date" style={inputStyle} value={endDate} onChange={(e) => setEndDate(e.target.value)} />
            </Field>
            <div style={{ fontFamily: sans, fontSize: 11.5, color: C.inkFaint, lineHeight: 1.5, marginTop: -4 }}>
              Reports and the dashboard keep this person's cost in every period up to and including their departure date, whether or not they're marked Active. Active/Inactive only controls whether they can still be assigned new work.
            </div>
            <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 12, marginTop: 2 }}>
              <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink, marginBottom: 10 }}>Record a new labor cost change</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <Field label="Gross salary">
                  <input type="number" min="0" style={inputStyle} value={grossSalary} onChange={(e) => setGrossSalary(e.target.value)} placeholder="e.g. 1400" />
                </Field>
                <Field label="Social security (employer)">
                  <input type="number" min="0" style={inputStyle} value={socialSecurity} onChange={(e) => setSocialSecurity(e.target.value)} placeholder="e.g. 320" />
                </Field>
                <Field label="Ticket restaurant">
                  <input type="number" min="0" style={inputStyle} value={ticketRestaurant} onChange={(e) => setTicketRestaurant(e.target.value)} placeholder="e.g. 88" />
                </Field>
                <Field label="Other insurance (optional)">
                  <input type="number" min="0" style={inputStyle} value={insurance} onChange={(e) => setInsurance(e.target.value)} placeholder="e.g. 0" />
                </Field>
                <Field label="Other cost">
                  <input type="number" min="0" style={inputStyle} value={otherCost} onChange={(e) => setOtherCost(e.target.value)} placeholder="Optional" />
                </Field>
                <Field label="Effective from">
                  <input type="date" style={inputStyle} value={costEffectiveDate} onChange={(e) => setCostEffectiveDate(e.target.value)} />
                </Field>
              </div>
              <div style={{ fontFamily: sans, fontSize: 11.5, color: C.inkFaint, marginTop: 8 }}>
                Cost per hour = total monthly cost ÷ (22 working days × daily hours from the contracted weekly hours).
                Saving with different numbers above adds a new row below, effective from that date — it never overwrites past periods.
              </div>
            </div>
            {liveEmployee && liveEmployee.costHistory && liveEmployee.costHistory.length > 0 && (
              <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 12 }}>
                <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink, marginBottom: 8 }}>Cost history</div>
                <div className="scrollbar-thin" style={{ maxHeight: 220, overflowY: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: sans, fontSize: 12 }}>
                    <thead>
                      <tr>
                        {["From", "To", "Gross", "SS", "Ticket", "Ins.", "Other", "Total", ""].map((h) => (
                          <th key={h} style={{ textAlign: "left", padding: "3px 6px", fontSize: 10.5, fontWeight: 600, color: C.inkFaint, borderBottom: `1px solid ${C.border}` }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {[...withPeriods(liveEmployee.costHistory)].reverse().map((h) => {
                        const total = (h.grossSalary || 0) + (h.socialSecurity || 0) + (h.ticketRestaurant || 0) + (h.insurance || 0) + (h.otherCost || 0);
                        const isEditing = editingHistoryId === h.id;
                        return isEditing ? (
                          <tr key={h.id}>
                            <td colSpan={9} style={{ padding: "6px" }}>
                              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                                <input type="date" style={{ ...inputStyle, width: 130, padding: "4px 6px" }} value={histForm.effectiveDate} onChange={(ev) => setHistForm((f) => ({ ...f, effectiveDate: ev.target.value }))} />
                                <input type="number" style={{ ...inputStyle, width: 70, padding: "4px 6px" }} value={histForm.grossSalary} onChange={(ev) => setHistForm((f) => ({ ...f, grossSalary: ev.target.value }))} placeholder="Gross" />
                                <input type="number" style={{ ...inputStyle, width: 70, padding: "4px 6px" }} value={histForm.socialSecurity} onChange={(ev) => setHistForm((f) => ({ ...f, socialSecurity: ev.target.value }))} placeholder="SS" />
                                <input type="number" style={{ ...inputStyle, width: 70, padding: "4px 6px" }} value={histForm.ticketRestaurant} onChange={(ev) => setHistForm((f) => ({ ...f, ticketRestaurant: ev.target.value }))} placeholder="Ticket" />
                                <input type="number" style={{ ...inputStyle, width: 60, padding: "4px 6px" }} value={histForm.insurance} onChange={(ev) => setHistForm((f) => ({ ...f, insurance: ev.target.value }))} placeholder="Ins." />
                                <input type="number" style={{ ...inputStyle, width: 60, padding: "4px 6px" }} value={histForm.otherCost} onChange={(ev) => setHistForm((f) => ({ ...f, otherCost: ev.target.value }))} placeholder="Other" />
                                <button onClick={() => saveHistoryEdit()} style={iconBtnStyle}><Check size={14} color={C.accent} /></button>
                                <button onClick={cancelEditHistory} style={iconBtnStyle}><X size={14} color={C.inkMuted} /></button>
                              </div>
                            </td>
                          </tr>
                        ) : (
                          <tr key={h.id}>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.from}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.to}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.grossSalary || 0}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.socialSecurity || 0}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.ticketRestaurant || 0}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.insurance || 0}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink }}>{h.otherCost || 0}</td>
                            <td style={{ padding: "3px 6px", fontFamily: mono, color: C.ink, fontWeight: 600 }}>{fmtEur(total)}</td>
                            <td style={{ padding: "3px 6px" }}>
                              <div style={{ display: "flex", gap: 4 }}>
                                <button onClick={() => startEditHistory(h)} style={iconBtnStyle}><Pencil size={12} color={C.inkMuted} /></button>
                                <button onClick={() => deleteHistoryEntry(h.id)} style={iconBtnStyle}><Trash2 size={12} color={C.inkMuted} /></button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 6 }}>
              <Btn variant="secondary" onClick={() => setModal(null)}>Cancel</Btn>
              <Btn onClick={save}>Save employee</Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}


/* ---------------------------------------------------------------------- */
/* Admin: Rental — properties and their expenses                            */
/* ---------------------------------------------------------------------- */

function AdminRental({ properties, expenses, refetchRental }) {
  const [viewPeriod, setViewPeriod] = useState("month");
  const [viewAnchor, setViewAnchor] = useState(TODAY);
  const [propModal, setPropModal] = useState(null); // null | "new" | property
  const [propName, setPropName] = useState("");
  const [propAddress, setPropAddress] = useState("");
  const [propNote, setPropNote] = useState("");
  const [expModal, setExpModal] = useState(false);
  const [expEditingId, setExpEditingId] = useState(null);
  const [expKind, setExpKind] = useState("recurring");
  const [expPropertyId, setExpPropertyId] = useState("");
  const [expMonth, setExpMonth] = useState(toKey(TODAY).slice(0, 7));
  const [expFromMonth, setExpFromMonth] = useState(toKey(TODAY).slice(0, 7));
  const [expToMonth, setExpToMonth] = useState("");
  const [expCategory, setExpCategory] = useState("rent");
  const [expAmount, setExpAmount] = useState("");
  const [expNote, setExpNote] = useState("");
  const [propFilter, setPropFilter] = useState("all");

  const [rangeStart, rangeEnd] = periodRange(viewPeriod, viewAnchor);
  const accrued = rentalExpensesForRange(expenses, rangeStart, rangeEnd);
  const propName_ = (id) => properties.find((p) => p.id === id)?.name || "—";

  // Entries that apply to at least one month of the period (for the detail list).
  const monthsInRange = [];
  for (let d = startOfMonth(rangeStart); d <= rangeEnd; d = new Date(d.getFullYear(), d.getMonth() + 1, 1)) monthsInRange.push(monthKeyOf(d));
  const startKeyOf = (x) => x.kind === "recurring" ? (x.fromMonth || "") : (x.month || "");
  const periodEntries = expenses
    .filter((x) => monthsInRange.some((m) => rentalExpenseAppliesTo(x, m)))
    .filter((x) => propFilter === "all" || x.propertyId === propFilter)
    .sort((a, b) => (a.kind === b.kind ? 0 : a.kind === "recurring" ? -1 : 1) || startKeyOf(b).localeCompare(startKeyOf(a)) || propName_(a.propertyId).localeCompare(propName_(b.propertyId)));
  const whenLabel = (x) => x.kind === "recurring" ? `${fmtMonthKey(x.fromMonth)} → ${x.toMonth ? fmtMonthKey(x.toMonth) : "open"}` : fmtMonthKey(x.month);

  function openNewProperty() { setPropModal("new"); setPropName(""); setPropAddress(""); setPropNote(""); }
  function openEditProperty(p) { setPropModal(p); setPropName(p.name); setPropAddress(p.address || ""); setPropNote(p.note || ""); }
  async function saveProperty() {
    if (!propName.trim()) return;
    const row = { name: propName.trim(), address: propAddress.trim() || null, note: propNote.trim() || null };
    const { error } = propModal === "new"
      ? await supabase.from("rental_properties").insert(row)
      : await supabase.from("rental_properties").update(row).eq("id", propModal.id);
    if (error) { window.alert(`Couldn't save property: ${error.message}`); return; }
    await refetchRental();
    setPropModal(null);
  }
  async function togglePropertyActive(p) {
    await supabase.from("rental_properties").update({ active: !p.active }).eq("id", p.id);
    refetchRental();
  }
  async function deleteProperty(p) {
    if (!window.confirm(`Permanently delete "${p.name}"? This cannot be undone.`)) return;
    const { error } = await supabase.from("rental_properties").delete().eq("id", p.id);
    if (error) {
      if (error.code === "23503") window.alert(`Can't delete "${p.name}" — it still has expenses booked. Delete those first, or deactivate the property instead.`);
      else window.alert(`Couldn't delete "${p.name}": ${error.message}`);
      return;
    }
    refetchRental();
  }

  function openNewExpense(propertyId, kind = "recurring") {
    setExpEditingId(null);
    setExpKind(kind);
    setExpPropertyId(propertyId || properties.find((p) => p.active)?.id || "");
    const m = monthKeyOf(rangeStart);
    setExpMonth(m); setExpFromMonth(m); setExpToMonth("");
    setExpCategory("rent"); setExpAmount(""); setExpNote("");
    setExpModal(true);
  }
  function openEditExpense(x) {
    setExpEditingId(x.id);
    setExpKind(x.kind);
    setExpPropertyId(x.propertyId);
    setExpMonth(x.month || monthKeyOf(rangeStart)); setExpFromMonth(x.fromMonth || monthKeyOf(rangeStart)); setExpToMonth(x.toMonth || "");
    setExpCategory(x.category); setExpAmount(String(x.amount)); setExpNote(x.note || "");
    setExpModal(true);
  }
  async function saveExpense() {
    if (!expPropertyId) { window.alert("Choose a property first."); return; }
    if (expKind === "recurring" && !(Number(expAmount) > 0)) { window.alert("Enter a monthly amount greater than zero."); return; }
    if (expKind === "extra" && !(Number(expAmount) !== 0 && !Number.isNaN(Number(expAmount)) && expAmount !== "")) { window.alert("Enter an amount (negative for a discount or rebate)."); return; }
    if (expKind === "extra" && !expMonth) { window.alert("Choose the month this extra expense belongs to."); return; }
    if (expKind === "recurring" && !expFromMonth) { window.alert("Choose the month the recurring expense starts."); return; }
    if (expKind === "recurring" && expToMonth && expToMonth < expFromMonth) { window.alert('"To" month must be the same as or after the "From" month.'); return; }
    const row = {
      property_id: expPropertyId, kind: expKind, category: expCategory, amount: Number(expAmount), note: expNote.trim() || null,
      month: expKind === "extra" ? expMonth : null,
      from_month: expKind === "recurring" ? expFromMonth : null,
      to_month: expKind === "recurring" ? (expToMonth || null) : null,
    };
    const { error } = expEditingId
      ? await supabase.from("rental_expenses").update(row).eq("id", expEditingId)
      : await supabase.from("rental_expenses").insert(row);
    if (error) { window.alert(`Couldn't save expense: ${error.message}`); return; }
    await refetchRental();
    setExpModal(false);
  }
  async function deleteExpense(id) {
    await supabase.from("rental_expenses").delete().eq("id", id);
    refetchRental();
  }

  function exportRental() {
    exportToExcel(`rental-expenses-${toKey(TODAY)}.xlsx`, [
      { name: "By property", rows: properties.map((p) => {
        const x = accrued.byProperty[p.id] || {};
        return { Property: p.name, Address: p.address || "", Status: p.active ? "Active" : "Inactive",
          ...Object.fromEntries(RENTAL_EXPENSE_CATEGORIES.map((k) => [`${RENTAL_EXPENSE_LABELS[k]} (€)`, Number((x[k] || 0).toFixed(2))])),
          "Total (€)": Number((x.total || 0).toFixed(2)) };
      }) },
      { name: "Entries", rows: periodEntries.map((x) => ({
        Type: x.kind === "recurring" ? "Recurring" : "Extra", From: x.kind === "recurring" ? x.fromMonth : x.month, To: x.kind === "recurring" ? (x.toMonth || "open") : x.month,
        Property: propName_(x.propertyId), Category: RENTAL_EXPENSE_LABELS[x.category] || x.category, "Amount (€/month)": x.amount, Note: x.note,
      })) },
    ]);
  }

  return (
    <div>
      <PageHeader title="Rental" sub="Rental properties and their expenses — rent, utilities and common expenses, booked per property and month."
        right={
          <div style={{ display: "flex", gap: 8 }}>
            <Btn variant="secondary" onClick={exportRental}>Export to Excel</Btn>
            <Btn variant="secondary" icon={Plus} onClick={openNewProperty}>Add property</Btn>
            <Btn variant="secondary" icon={Plus} onClick={() => openNewExpense(null, "extra")} disabled={properties.length === 0}>Add extra</Btn>
            <Btn icon={Plus} onClick={() => openNewExpense(null, "recurring")} disabled={properties.length === 0}>Add recurring expense</Btn>
          </div>
        } />
      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 10 }}>
          Recurring expenses (rent, standing charges) run from a start month to an end month or open-ended; extras are one-off amounts for a single month. Everything is accrued to this reporting period (each month's amount is spread evenly over its days, so weekly and quarterly views reconcile with the dashboard).
        </div>
        <PeriodSelector period={viewPeriod} setPeriod={setViewPeriod} anchor={viewAnchor} setAnchor={setViewAnchor} />
      </Panel>

      <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
        {RENTAL_EXPENSE_CATEGORIES.map((k) => (
          <StatCard key={k} label={RENTAL_EXPENSE_LABELS[k]} value={fmtEur(accrued.byCategory[k])} sub="This period" />
        ))}
        <StatCard label="Total rental expenses" value={fmtEur(accrued.total)} accent={C.accent} sub="Goes to the Rental cost center" />
      </div>

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 12 }}>Expenses by property</div>
        {properties.length === 0 ? (
          <EmptyState icon={Home} title="No rental properties yet" sub="Add a property to start booking its expenses." />
        ) : (
          <TableShell headers={["Property", "Address", ...RENTAL_EXPENSE_CATEGORIES.map((k) => RENTAL_EXPENSE_LABELS[k]), "Total", "Status", ""]}>
            {properties.map((p) => {
              const x = accrued.byProperty[p.id] || {};
              return (
                <tr key={p.id} style={{ opacity: p.active ? 1 : 0.55 }}>
                  <Td title={p.note || undefined}>{p.name}</Td>
                  <Td style={{ color: C.inkMuted }}>{p.address || "—"}</Td>
                  {RENTAL_EXPENSE_CATEGORIES.map((k) => <Td key={k} mono>{Math.abs(x[k] || 0) > 0.005 ? fmtEur(x[k]) : "—"}</Td>)}
                  <Td mono style={{ fontWeight: 700 }}>{Math.abs(x.total || 0) > 0.005 ? fmtEur(x.total) : "—"}</Td>
                  <Td>{p.active ? <Badge tone="accent">Active</Badge> : <Badge tone="neutral">Inactive</Badge>}</Td>
                  <Td right>
                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      <Btn size="sm" variant="ghost" icon={Plus} onClick={() => openNewExpense(p.id)}>Expense</Btn>
                      <Btn size="sm" variant="ghost" icon={Pencil} onClick={() => openEditProperty(p)}>Edit</Btn>
                      <Btn size="sm" variant={p.active ? "danger" : "secondary"} onClick={() => togglePropertyActive(p)}>{p.active ? "Deactivate" : "Reactivate"}</Btn>
                      <Btn size="sm" variant="ghost" icon={Trash2} onClick={() => deleteProperty(p)}>Delete</Btn>
                    </div>
                  </Td>
                </tr>
              );
            })}
            <tr style={{ fontWeight: 700 }}>
              <Td>Total</Td>
              <Td />
              {RENTAL_EXPENSE_CATEGORIES.map((k) => <Td key={k} mono>{fmtEur(accrued.byCategory[k])}</Td>)}
              <Td mono>{fmtEur(accrued.total)}</Td>
              <Td /><Td />
            </tr>
          </TableShell>
        )}
      </Panel>

      <Panel>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink }}>Expenses applying to this period</div>
          <select style={{ ...inputStyle, width: 220 }} value={propFilter} onChange={(e) => setPropFilter(e.target.value)}>
            <option value="all">All properties</option>
            {properties.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
          </select>
        </div>
        {periodEntries.length === 0 ? (
          <EmptyState icon={Euro} title="No expenses booked for these months" />
        ) : (
          <TableShell headers={["Type", "Period", "Property", "Category", "Amount / month", "Note", ""]}>
            {periodEntries.map((x) => (
              <tr key={x.id}>
                <Td><Badge tone={x.kind === "recurring" ? "accent" : "warn"}>{x.kind === "recurring" ? "Recurring" : "Extra"}</Badge></Td>
                <Td mono>{whenLabel(x)}</Td>
                <Td>{propName_(x.propertyId)}</Td>
                <Td>{RENTAL_EXPENSE_LABELS[x.category] || x.category}</Td>
                <Td mono style={{ color: x.amount < 0 ? C.accentDark : undefined }}>{fmtEur(x.amount)}{x.amount < 0 ? " (discount)" : ""}</Td>
                <Td style={{ color: C.inkMuted }}>{x.note || "—"}</Td>
                <Td right>
                  <div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}>
                    <button onClick={() => openEditExpense(x)} style={iconBtnStyle} title="Edit"><Pencil size={13} color={C.inkMuted} /></button>
                    <button onClick={() => deleteExpense(x.id)} style={iconBtnStyle} title="Delete"><Trash2 size={13} color={C.inkMuted} /></button>
                  </div>
                </Td>
              </tr>
            ))}
          </TableShell>
        )}
      </Panel>

      {propModal && (
        <Modal title={propModal === "new" ? "Add rental property" : "Edit rental property"} onClose={() => setPropModal(null)} width={440}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Field label="Property name">
              <input style={inputStyle} value={propName} onChange={(e) => setPropName(e.target.value)} placeholder="e.g. Kaniggos 2 — 3rd floor" autoFocus />
            </Field>
            <Field label="Address (optional)">
              <input style={inputStyle} value={propAddress} onChange={(e) => setPropAddress(e.target.value)} placeholder="Street, number, city" />
            </Field>
            <Field label="Note (optional)">
              <input style={inputStyle} value={propNote} onChange={(e) => setPropNote(e.target.value)} placeholder="e.g. Tenant, lease dates" />
            </Field>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 6 }}>
              <Btn variant="secondary" onClick={() => setPropModal(null)}>Cancel</Btn>
              <Btn onClick={saveProperty}>Save property</Btn>
            </div>
          </div>
        </Modal>
      )}

      {expModal && (
        <Modal title={expEditingId ? "Edit rental expense" : "Add rental expense"} onClose={() => setExpModal(false)} width={420}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <div style={{ display: "flex", border: `1px solid ${C.border}`, borderRadius: 7, overflow: "hidden" }}>
              {Object.keys(RENTAL_EXPENSE_KINDS).map((k) => (
                <button key={k} type="button" onClick={() => setExpKind(k)}
                  style={{ flex: 1, padding: "8px 10px", fontSize: 12.5, fontFamily: sans, fontWeight: 600, border: "none", cursor: "pointer",
                    background: expKind === k ? C.accent : C.surface, color: expKind === k ? "#fff" : C.inkMuted }}>
                  {RENTAL_EXPENSE_KINDS[k]}
                </button>
              ))}
            </div>
            <Field label="Property">
              <select style={inputStyle} value={expPropertyId} onChange={(e) => setExpPropertyId(e.target.value)}>
                <option value="">— choose —</option>
                {properties.filter((p) => p.active || p.id === expPropertyId).map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
            </Field>
            {expKind === "extra" ? (
              <Field label="Month">
                <input type="month" style={inputStyle} value={expMonth} onChange={(e) => setExpMonth(e.target.value)} />
              </Field>
            ) : (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <Field label="From month">
                  <input type="month" style={inputStyle} value={expFromMonth} onChange={(e) => setExpFromMonth(e.target.value)} />
                </Field>
                <Field label="To month (leave empty = open-ended)">
                  <input type="month" style={inputStyle} value={expToMonth} onChange={(e) => setExpToMonth(e.target.value)} />
                </Field>
              </div>
            )}
            <Field label="Category">
              <select style={inputStyle} value={expCategory} onChange={(e) => setExpCategory(e.target.value)}>
                {RENTAL_EXPENSE_CATEGORIES.map((k) => <option key={k} value={k}>{RENTAL_EXPENSE_LABELS[k]}</option>)}
              </select>
            </Field>
            <Field label={expKind === "recurring" ? "Amount per month (€)" : "Amount (€) — enter a negative amount for a discount or rebate"}>
              <input type="number" min={expKind === "recurring" ? "0" : undefined} style={inputStyle} value={expAmount} onChange={(e) => setExpAmount(e.target.value)} placeholder={expKind === "recurring" ? "e.g. 850" : "e.g. 120 or -300"} autoFocus />
            </Field>
            <Field label="Note (optional)">
              <input style={inputStyle} value={expNote} onChange={(e) => setExpNote(e.target.value)} placeholder="e.g. DEH electricity bill" />
            </Field>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 6 }}>
              <Btn variant="secondary" onClick={() => setExpModal(false)}>Cancel</Btn>
              <Btn onClick={saveExpense}>{expEditingId ? "Save changes" : "Add expense"}</Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}


/* ---------------------------------------------------------------------- */
/* Admin: Report builder — pick a dimension, get revenue/cost/profit      */
/* ---------------------------------------------------------------------- */

const REPORT_DIMENSIONS = [
  { key: "client", label: "Client" },
  { key: "employee", label: "Employee" },
  { key: "costCenter", label: "Cost center" },
  { key: "department", label: "Department" },
  { key: "month", label: "Month" },
  { key: "category", label: "Category (chargeable work)" },
  { key: "activity", label: "Activity (non-chargeable)" },
];
const REPORT_METRICS = [
  { key: "hours", label: "Hours", format: "hours" },
  { key: "revenue", label: "Revenue", format: "eur" },
  { key: "cost", label: "Cost", format: "eur" },
  { key: "profit", label: "Profit", format: "eur" },
  { key: "margin", label: "Margin %", format: "pct" },
];
// Which metrics are meaningful for each dimension — revenue/cost/profit/
// margin aren't tracked per activity (non-chargeable, by definition), and
// hours aren't meaningful for Director/Finance/Rental cost centers (nobody
// logs hours there).
const REPORT_METRICS_FOR = {
  client: ["hours", "revenue", "cost", "profit", "margin"],
  employee: ["hours", "revenue", "cost", "profit", "margin"],
  costCenter: ["hours", "revenue", "cost", "profit", "margin"],
  department: ["hours", "revenue", "cost", "profit", "margin"],
  month: ["hours", "revenue", "cost", "profit", "margin"],
  category: ["hours", "revenue", "cost", "profit", "margin"],
  activity: ["hours"],
};

function AdminReportBuilder({ employees, clients, entries, rentalExpenses }) {
  const [dimension, setDimension] = useState("client");
  const [metric, setMetric] = useState("profit");
  const [period, setPeriod] = useState("quarter");
  const [anchor, setAnchor] = useState(TODAY);
  const [customFrom, setCustomFrom] = useState("");
  const [customTo, setCustomTo] = useState("");
  const [clientFilter, setClientFilter] = useState("all");
  const [employeeFilter, setEmployeeFilter] = useState("all");
  const [departmentFilter, setDepartmentFilter] = useState("all");
  const [sortDir, setSortDir] = useState("desc");
  const [showChart, setShowChart] = useState(true);
  const [compareTo, setCompareTo] = useState("none"); // "none" | "previous" | "yoy" — ignored for the Month dimension, which always shows YoY per row

  const usingCustomRange = !!(customFrom && customTo);
  const [presetStart, presetEnd] = periodRange(period, anchor);
  const rangeStart = usingCustomRange ? fromKey(customFrom) : presetStart;
  const rangeEnd = usingCustomRange ? fromKey(customTo) : presetEnd;
  const validCustomRange = !usingCustomRange || customFrom <= customTo;

  // Apply the three filters up front, to every downstream computation —
  // this is what makes the "same" dimension report answer a narrower
  // question (e.g. Cost center report, Accounting Team only).
  const filteredClients = clients.filter((c) => clientFilter === "all" || c.id === clientFilter);
  const filteredEmployees = employees.filter((e) =>
    (departmentFilter === "all" || (e.department || DEFAULT_DEPARTMENT) === departmentFilter) &&
    (employeeFilter === "all" || e.id === employeeFilter));
  const includeRental = clientFilter === "all" && employeeFilter === "all" && (departmentFilter === "all" || departmentFilter === "rental");
  const scopedEntries = entries.filter((e) =>
    (!e.clientId || filteredClients.some((c) => c.id === e.clientId)) &&
    filteredEmployees.some((emp) => emp.id === e.employeeId));

  // Core aggregate for one arbitrary date range, given the (already
  // filtered) clients/employees/entries in scope. Mirrors the dashboard's
  // computation exactly, so figures reconcile with it when filters are
  // left at "All".
  function computeCore(rStart, rEnd) {
    const rangeEntries = scopedEntries.filter((e) => { const d = fromKey(e.date); return d >= rStart && d <= rEnd; });
    const employeeCosts = filteredEmployees
      .filter((e) => employeeBilledInRange(e, rStart, rEnd))
      .map((emp) => {
        const empEntries = rangeEntries.filter((e) => e.employeeId === emp.id);
        return { id: emp.id, department: emp.department || DEFAULT_DEPARTMENT, entries: empEntries, hours: empEntries.reduce((s, e) => s + e.hours, 0), cost: periodEmployeeCost(emp, rStart, rEnd) };
      });
    const byClient = filteredClients
      .filter((c) => {
        const start = earliestHistoryDate(c.feeHistory);
        if (start && fromKey(start) > rEnd) return false;
        if (c.endDate && fromKey(c.endDate) < rStart) return false;
        return true;
      })
      .map((c) => {
        const clientEntries = rangeEntries.filter((e) => e.clientId === c.id);
        const hrs = clientEntries.reduce((s, e) => s + e.hours, 0);
        const { revenue, revenueByCostCenter } = periodClientMetrics(c, clientEntries, rStart, rEnd);
        const cost = employeeCosts.reduce((sum, ec) => {
          if (ec.hours <= 0) return sum;
          const empHrsForClient = clientEntries.filter((e) => e.employeeId === ec.id).reduce((s, e) => s + e.hours, 0);
          return sum + ec.cost * (empHrsForClient / ec.hours);
        }, 0);
        const profit = revenue - cost;
        return { id: c.id, label: c.name, hours: hrs, revenue, revenueByCostCenter, cost, profit, margin: revenue > 0 ? (profit / revenue) * 100 : null };
      });
    const totalRevenue = byClient.reduce((s, c) => s + c.revenue, 0);
    const byEmployee = filteredEmployees
      .filter((e) => employeeBilledInRange(e, rStart, rEnd))
      .map((emp) => {
        const hrs = rangeEntries.filter((e) => e.employeeId === emp.id).reduce((s, e) => s + e.hours, 0);
        const revenue = byClient.reduce((sum, c) => {
          if (c.hours <= 0) return sum;
          const empHrsForClient = rangeEntries.filter((e) => e.employeeId === emp.id && e.clientId === c.id).reduce((s, e) => s + e.hours, 0);
          return sum + c.revenue * (empHrsForClient / c.hours);
        }, 0);
        const cost = employeeCosts.find((ec) => ec.id === emp.id)?.cost || 0;
        const profit = revenue - cost;
        return { id: emp.id, label: emp.name, hours: hrs, revenue, cost, profit, margin: revenue > 0 ? (profit / revenue) * 100 : null };
      });
    const rentalExp = includeRental ? rentalExpensesForRange(rentalExpenses, rStart, rEnd) : { total: 0, byCategory: emptyByCC() };
    const totalLaborCost = byEmployee.reduce((s, e) => s + e.cost, 0);
    const totalCost = totalLaborCost + rentalExp.total;
    const revenueByCC = emptyByCC();
    byClient.forEach((c) => { Object.keys(revenueByCC).forEach((cc) => { revenueByCC[cc] += c.revenueByCostCenter?.[cc] || 0; }); });
    employeeCosts.forEach((ec) => {
      const emp = filteredEmployees.find((e) => e.id === ec.id);
      ec.byCostCenter = employeeCostByCostCenter(emp, ec.cost, ec.entries, revenueByCC);
    });
    const costByDepartment = Object.fromEntries(DEPARTMENTS.map((d) => [d.key, 0]));
    costByDepartment.rental += rentalExp.total;
    employeeCosts.forEach((ec) => { costByDepartment[ec.department] = (costByDepartment[ec.department] || 0) + ec.cost; });
    const byCostCenter = [...COST_CENTERS, UNALLOCATED].map((cc) => {
      const revenue = revenueByCC[cc] || 0;
      const cost = employeeCosts.reduce((s, ec) => s + (ec.byCostCenter?.[cc] || 0), 0) + (cc === "rental" ? rentalExp.total : 0);
      const hours = CATEGORIES.includes(cc) ? rangeEntries.filter((e) => !e.activity && e.category === cc).reduce((s, e) => s + e.hours, 0) : null;
      const profit = revenue - cost;
      return { key: cc, label: costCenterLabel(cc), hours, revenue, cost, profit, margin: revenue > 0 ? (profit / revenue) * 100 : null };
    });
    const byDepartment = DEPARTMENTS.map((dep) => {
      const rows = byCostCenter.filter((r) => dep.costCenters.includes(r.key));
      const revenue = rows.reduce((s, r) => s + r.revenue, 0);
      const cost = costByDepartment[dep.key] || 0;
      const hours = employeeCosts.filter((ec) => ec.department === dep.key).reduce((s, ec) => s + ec.hours, 0);
      const profit = revenue - cost;
      return { key: dep.key, label: departmentLabel(dep.key), hours, revenue, cost, profit, margin: revenue > 0 ? (profit / revenue) * 100 : null };
    });
    const totalHoursInPeriod = rangeEntries.reduce((s, e) => s + e.hours, 0);
    const byActivity = ACTIVITIES.map((a) => ({ key: a, label: ACTIVITY_LABELS[a], hours: rangeEntries.filter((e) => e.activity === a).reduce((s, e) => s + e.hours, 0), revenue: 0, cost: 0, profit: null, margin: null }));
    return { byClient, byEmployee, byCostCenter, byDepartment, byActivity, totalRevenue, totalCost, totalHoursInPeriod, unallocatedRevenue: revenueByCC[UNALLOCATED] || 0 };
  }

  // Build the row set for the chosen dimension, for an arbitrary range —
  // called once for the period on screen, and again (same dimension, a
  // shifted range) whenever a comparison is requested, so the two row sets
  // line up row-for-row by key.
  function buildRows(dim, rStart, rEnd) {
    const c = computeCore(rStart, rEnd);
    if (dim === "client") return c.byClient;
    if (dim === "employee") return c.byEmployee;
    if (dim === "costCenter") return c.byCostCenter.filter((r) => r.key !== UNALLOCATED || r.revenue > 0.005 || r.cost > 0.005);
    if (dim === "department") return c.byDepartment;
    if (dim === "activity") return c.byActivity;
    if (dim === "category") return c.byCostCenter.filter((r) => CATEGORIES.includes(r.key)).map((r) => ({ ...r, label: CATEGORY_LABELS[r.key] }));
    if (dim === "month") {
      const months = [];
      for (let d = startOfMonth(rStart); d <= rEnd; d = new Date(d.getFullYear(), d.getMonth() + 1, 1)) {
        const mStart = d > rStart ? d : rStart;
        const mEnd = endOfMonth(d) < rEnd ? endOfMonth(d) : rEnd;
        const mc = computeCore(mStart, mEnd);
        const hours = scopedEntries.filter((e) => { const dt = fromKey(e.date); return dt >= mStart && dt <= mEnd; }).reduce((s, e) => s + e.hours, 0);
        const profit = mc.totalRevenue - mc.totalCost;
        months.push({ key: toKey(d).slice(0, 7), label: d.toLocaleDateString("en-GB", { month: "short", year: "numeric" }), hours, revenue: mc.totalRevenue, cost: mc.totalCost, profit, margin: mc.totalRevenue > 0 ? (profit / mc.totalRevenue) * 100 : null });
      }
      return months;
    }
    return [];
  }
  const core = computeCore(rangeStart, rangeEnd);
  let rows = buildRows(dimension, rangeStart, rangeEnd);
  const availableMetrics = REPORT_METRICS_FOR[dimension];
  const effectiveMetric = availableMetrics.includes(metric) ? metric : availableMetrics[0];

  // Growth: compare every row to the same dimension over the previous
  // period of equal length, or the same period a year ago. Month rows are
  // always compared to the same calendar month last year, regardless of
  // this toggle, since each row is already its own period.
  const { prevStart: cmpPrevStart, prevEnd: cmpPrevEnd, yoyStart: cmpYoyStart, yoyEnd: cmpYoyEnd } = comparisonRanges(rangeStart, rangeEnd);
  const rowKey = (r) => r.key ?? r.id;
  if (dimension === "month") {
    rows = rows.map((r) => {
      const d = fromKey(r.key + "-01");
      const yStart = new Date(d.getFullYear() - 1, d.getMonth(), 1);
      const yEnd = endOfMonth(yStart);
      const yc = computeCore(yStart, yEnd);
      const yoyVal = effectiveMetric === "hours" ? scopedEntries.filter((e) => { const dt = fromKey(e.date); return dt >= yStart && dt <= yEnd; }).reduce((s, e) => s + e.hours, 0) : (effectiveMetric === "cost" ? yc.totalCost : effectiveMetric === "profit" ? yc.totalRevenue - yc.totalCost : effectiveMetric === "margin" ? (yc.totalRevenue > 0 ? ((yc.totalRevenue - yc.totalCost) / yc.totalRevenue) * 100 : null) : yc.totalRevenue);
      return { ...r, compareValue: yoyVal, growthPct: pctChange(r[effectiveMetric], yoyVal) };
    });
  } else if (compareTo !== "none") {
    const [cStart, cEnd] = compareTo === "previous" ? [cmpPrevStart, cmpPrevEnd] : [cmpYoyStart, cmpYoyEnd];
    const cmpRows = buildRows(dimension, cStart, cEnd);
    const byKey = Object.fromEntries(cmpRows.map((r) => [rowKey(r), r]));
    rows = rows.map((r) => {
      const match = byKey[rowKey(r)];
      const compareValue = match ? match[effectiveMetric] : (dimension === "client" || dimension === "employee" ? 0 : null);
      return { ...r, compareValue, growthPct: compareValue !== null && compareValue !== undefined ? pctChange(r[effectiveMetric], compareValue) : null };
    });
  }

  const unallocatedNote = ["client", "costCenter", "department", "month"].includes(dimension) && core.unallocatedRevenue > 0.005;

  const sortedRows = [...rows].sort((a, b) => {
    const av = a[effectiveMetric] ?? -Infinity, bv = b[effectiveMetric] ?? -Infinity;
    return sortDir === "desc" ? bv - av : av - bv;
  });
  const totalRow = dimension !== "month" && dimension !== "activity" ? {
    key: "total", label: "Total", hours: rows.reduce((s, r) => s + (r.hours || 0), 0),
    revenue: core.totalRevenue, cost: core.totalCost, profit: core.totalRevenue - core.totalCost,
    margin: core.totalRevenue > 0 ? ((core.totalRevenue - core.totalCost) / core.totalRevenue) * 100 : null,
  } : null;

  const fmtMetric = (v, fmt) => {
    if (v === null || v === undefined) return "—";
    if (fmt === "eur") return fmtEur(v);
    if (fmt === "pct") return v.toFixed(1) + "%";
    return v.toFixed(2) + "h";
  };
  const showGrowthColumn = dimension === "month" || compareTo !== "none";
  const metricDef = REPORT_METRICS.find((m) => m.key === effectiveMetric);
  const chartData = sortedRows.slice(0, 20).map((r) => ({ name: r.label.length > 16 ? r.label.slice(0, 15) + "…" : r.label, [metricDef.label]: Number((r[effectiveMetric] || 0).toFixed(2)) }));

  function exportReport() {
    const cols = availableMetrics.map((mk) => REPORT_METRICS.find((m) => m.key === mk));
    const dimLabel = REPORT_DIMENSIONS.find((d) => d.key === dimension).label;
    exportToExcel(`report-${dimension}-${toKey(TODAY)}.xlsx`, [
      { name: "Report", rows: [...sortedRows, ...(totalRow ? [totalRow] : [])].map((r) => ({
        [dimLabel]: r.label,
        ...Object.fromEntries(cols.map((c) => [`${c.label} (${c.format === "eur" ? "€" : c.format === "pct" ? "%" : "h"})`, r[c.key] === null || r[c.key] === undefined ? "" : Number(r[c.key].toFixed(2))])),
        ...(showGrowthColumn ? { [`Compare (${metricDef.label})`]: r.compareValue === null || r.compareValue === undefined ? "" : Number(r.compareValue.toFixed(2)), "Growth (%)": r.growthPct === null || r.growthPct === undefined ? "" : Number(r.growthPct.toFixed(1)) } : {}),
      })) },
    ]);
  }

  return (
    <div>
      <PageHeader title="Reports" sub="Pick a dimension and a metric to build the table you need, from the whole database."
        right={<Btn variant="secondary" onClick={exportReport}>Export to Excel</Btn>} />

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
          <Field label="Group by">
            <select style={{ ...inputStyle, width: 220 }} value={dimension} onChange={(e) => setDimension(e.target.value)}>
              {REPORT_DIMENSIONS.map((d) => <option key={d.key} value={d.key}>{d.label}</option>)}
            </select>
          </Field>
          <Field label="Metric">
            <select style={{ ...inputStyle, width: 160 }} value={effectiveMetric} onChange={(e) => setMetric(e.target.value)}>
              {REPORT_METRICS.filter((m) => availableMetrics.includes(m.key)).map((m) => <option key={m.key} value={m.key}>{m.label}</option>)}
            </select>
          </Field>
          <Field label="Sort">
            <select style={{ ...inputStyle, width: 130 }} value={sortDir} onChange={(e) => setSortDir(e.target.value)}>
              <option value="desc">Highest first</option>
              <option value="asc">Lowest first</option>
            </select>
          </Field>
          <Field label="Chart">
            <select style={{ ...inputStyle, width: 100 }} value={showChart ? "on" : "off"} onChange={(e) => setShowChart(e.target.value === "on")}>
              <option value="on">Show</option>
              <option value="off">Hide</option>
            </select>
          </Field>
          <Field label={dimension === "month" ? "Growth (per row)" : "Compare to"}>
            {dimension === "month" ? (
              <div style={{ ...inputStyle, width: 210, color: C.inkMuted, background: C.surfaceMuted || C.surface }}>Same month, last year</div>
            ) : (
              <select style={{ ...inputStyle, width: 210 }} value={compareTo} onChange={(e) => setCompareTo(e.target.value)}>
                <option value="none">No comparison</option>
                <option value="previous">Previous period (same length)</option>
                <option value="yoy">Same period, last year</option>
              </select>
            )}
          </Field>
        </div>

        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end", paddingTop: 12, borderTop: `1px solid ${C.border}` }}>
          <div style={{ opacity: usingCustomRange ? 0.45 : 1, pointerEvents: usingCustomRange ? "none" : "auto" }}>
            <PeriodSelector period={period} setPeriod={setPeriod} anchor={anchor} setAnchor={setAnchor} />
          </div>
          <Field label="Or custom range — from">
            <input type="date" style={{ ...inputStyle, width: 150 }} value={customFrom} onChange={(e) => setCustomFrom(e.target.value)} />
          </Field>
          <Field label="to">
            <input type="date" style={{ ...inputStyle, width: 150 }} value={customTo} onChange={(e) => setCustomTo(e.target.value)} />
          </Field>
          {usingCustomRange && <Btn variant="ghost" size="sm" onClick={() => { setCustomFrom(""); setCustomTo(""); }}>Clear custom range</Btn>}
        </div>
        {!validCustomRange && <div style={{ fontFamily: sans, fontSize: 12, color: C.danger, marginTop: 8 }}>"To" must be the same as or after "From".</div>}

        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", paddingTop: 12, marginTop: 12, borderTop: `1px solid ${C.border}` }}>
          <Field label="Client filter">
            <select style={{ ...inputStyle, width: 190 }} value={clientFilter} onChange={(e) => setClientFilter(e.target.value)}>
              <option value="all">All clients</option>
              {clients.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
          <Field label="Employee filter">
            <select style={{ ...inputStyle, width: 190 }} value={employeeFilter} onChange={(e) => setEmployeeFilter(e.target.value)}>
              <option value="all">All employees</option>
              {employees.map((e) => <option key={e.id} value={e.id}>{e.name}</option>)}
            </select>
          </Field>
          <Field label="Department filter">
            <select style={{ ...inputStyle, width: 190 }} value={departmentFilter} onChange={(e) => setDepartmentFilter(e.target.value)}>
              <option value="all">All departments</option>
              {DEPARTMENTS.map((d) => <option key={d.key} value={d.key}>{d.label}</option>)}
            </select>
          </Field>
        </div>
      </Panel>

      {rows.length === 0 ? (
        <EmptyState icon={SlidersHorizontal} title="No data for this combination" sub="Try a wider period or fewer filters." />
      ) : (
        <>
          {showChart && (
            <Panel style={{ marginBottom: 18 }}>
              <div style={{ width: "100%", height: 260 }}>
                <ResponsiveContainer>
                  <BarChart data={chartData} margin={{ left: -10, right: 10 }}>
                    <CartesianGrid stroke={C.border} vertical={false} />
                    <XAxis dataKey="name" tick={{ fontFamily: sans, fontSize: 10.5, fill: C.inkMuted }} axisLine={{ stroke: C.border }} tickLine={false} interval={0} angle={sortedRows.length > 8 ? -30 : 0} textAnchor={sortedRows.length > 8 ? "end" : "middle"} height={sortedRows.length > 8 ? 55 : 30} />
                    <YAxis tick={{ fontFamily: mono, fontSize: 11, fill: C.inkMuted }} axisLine={false} tickLine={false} />
                    <Tooltip contentStyle={{ fontFamily: sans, fontSize: 12.5, borderRadius: 8, border: `1px solid ${C.border}` }} formatter={(v) => fmtMetric(v, metricDef.format)} />
                    <Bar dataKey={metricDef.label} fill={C.accent} radius={[4, 4, 0, 0]} maxBarSize={34} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </Panel>
          )}
          <Panel>
            {unallocatedNote && (
              <div style={{ fontFamily: sans, fontSize: 12, color: C.warn, marginBottom: 10 }}>
                Some fixed fees have no cost-center split yet — their revenue ({fmtEur(core.unallocatedRevenue)}) is included in the total but not broken out below.
              </div>
            )}
            <TableShell headers={[REPORT_DIMENSIONS.find((d) => d.key === dimension).label, ...availableMetrics.map((mk) => REPORT_METRICS.find((m) => m.key === mk).label), ...(showGrowthColumn ? [dimension === "month" ? `${metricDef.label}, same month last year` : `${metricDef.label}, ${compareTo === "yoy" ? "same period last year" : "previous period"}`, "Growth"] : [])]}>
              {sortedRows.map((r) => (
                <tr key={r.key || r.id}>
                  <Td>{r.label}</Td>
                  {availableMetrics.map((mk) => {
                    const v = r[mk];
                    const fmt = REPORT_METRICS.find((m) => m.key === mk).format;
                    const negative = (mk === "profit" || mk === "margin") && typeof v === "number" && v < 0;
                    return <Td key={mk} mono style={{ color: negative ? C.danger : undefined, fontWeight: mk === effectiveMetric ? 700 : 400 }}>{fmtMetric(v, fmt)}</Td>;
                  })}
                  {showGrowthColumn && (
                    <>
                      <Td mono style={{ color: C.inkMuted }}>{fmtMetric(r.compareValue, metricDef.format)}</Td>
                      <Td mono style={{ color: r.growthPct === null ? undefined : r.growthPct < 0 ? C.danger : C.accentDark, fontWeight: 700 }}>
                        {r.growthPct === null ? "—" : `${r.growthPct > 0 ? "+" : ""}${r.growthPct.toFixed(1)}%`}
                      </Td>
                    </>
                  )}
                </tr>
              ))}
              {totalRow && (
                <tr style={{ fontWeight: 700, borderTop: `2px solid ${C.borderStrong}` }}>
                  <Td>Total</Td>
                  {availableMetrics.map((mk) => <Td key={mk} mono>{fmtMetric(totalRow[mk], REPORT_METRICS.find((m) => m.key === mk).format)}</Td>)}
                  {showGrowthColumn && <><Td /><Td /></>}
                </tr>
              )}
            </TableShell>
          </Panel>
        </>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Root                                                                     */
/* ---------------------------------------------------------------------- */

export default function App() {
  const [session, setSession] = useState(undefined);
  const [view, setView] = useState(null);
  const [passwordRecovery, setPasswordRecovery] = useState(false);

  useEffect(() => {
    if (!supabaseConfigured) { setSession(null); return; }
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((event, sess) => {
      if (event === "PASSWORD_RECOVERY") setPasswordRecovery(true);
      setSession(sess);
    });
    return () => sub.subscription.unsubscribe();
  }, []);

  const [profiles, refetchProfilesRaw, profilesLoading] = useTable("profiles", mapProfile, "name");
  const [costHistoryRows, refetchCostHistoryRaw, costHistoryLoading] = useTable("employee_cost_history", mapCostHistory);
  const [clientRows, refetchClientsRaw, clientsLoading] = useTable("clients", mapClientRow, "name");
  const [feeHistoryRows, refetchFeeHistoryRaw, feeHistoryLoading] = useTable("client_fee_history", mapFeeHistory);
  const [extraFeeRows, refetchExtraFeesRaw, extraFeesLoading] = useTable("client_extra_fees", mapExtraFee);
  const [entries, refetchEntries, entriesLoading] = useTable("time_entries", mapEntry);
  const [lockedWeeks, refetchLockedWeeks, lockedLoading] = useTable("locked_weeks", mapLockedWeek);
  const [leaveRequests, refetchLeaveRequests, leaveLoading] = useTable("leave_requests", mapLeaveRequest);
  const [rentalProperties, refetchRentalPropertiesRaw, rentalPropsLoading] = useTable("rental_properties", mapRentalProperty, "name");
  const [rentalExpenses, refetchRentalExpensesRaw, rentalExpLoading] = useTable("rental_expenses", mapRentalExpense);
  const refetchRental = useCallback(async () => { await Promise.all([refetchRentalPropertiesRaw(), refetchRentalExpensesRaw()]); }, [refetchRentalPropertiesRaw, refetchRentalExpensesRaw]);

  const refetchEmployees = useCallback(async () => { await Promise.all([refetchProfilesRaw(), refetchCostHistoryRaw()]); }, [refetchProfilesRaw, refetchCostHistoryRaw]);
  const refetchClients = useCallback(async () => { await Promise.all([refetchClientsRaw(), refetchFeeHistoryRaw(), refetchExtraFeesRaw()]); }, [refetchClientsRaw, refetchFeeHistoryRaw, refetchExtraFeesRaw]);

  const employees = React.useMemo(() => {
    return profiles.map((p) => {
      const history = costHistoryRows.filter((h) => h.employeeId === p.id);
      const latest = latestFromHistory(history, ["grossSalary", "socialSecurity", "ticketRestaurant", "insurance", "otherCost"])
        || { grossSalary: 0, socialSecurity: 0, ticketRestaurant: 0, insurance: 0, otherCost: 0 };
      return { ...p, ...latest, costHistory: history };
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [profiles, costHistoryRows]);

  const clients = React.useMemo(() => {
    return clientRows.map((c) => {
      const history = feeHistoryRows.filter((h) => h.clientId === c.id);
      const latest = latestFromHistory(history, ["fixedFee", "allocation", "feeSplit", "feeSplitNote"])
        || { fixedFee: c.fixedFee, allocation: { accounting: 0, tax: 0, payroll: 0, other: 0 }, feeSplit: emptyFeeSplit(), feeSplitNote: "" };
      const extraFeeEntries = extraFeeRows.filter((x) => x.clientId === c.id);
      return { ...c, ...latest, feeHistory: history, extraFeeEntries };
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [clientRows, feeHistoryRows, extraFeeRows]);

  const user = session ? employees.find((p) => p.id === session.user.id) : null;

  useEffect(() => {
    if (user && view === null) setView(user.role === "admin" ? "dashboard" : "timesheet");
  }, [user, view]);

  useEffect(() => {
    if (session && !profilesLoading && !user) {
      // Profile row not created yet right after sign-up — try again shortly.
      const t = setTimeout(() => refetchProfilesRaw(), 1200);
      return () => clearTimeout(t);
    }
  }, [session, profilesLoading, user, refetchProfilesRaw]);

  if (!supabaseConfigured) return <SetupScreen />;
  if (session === undefined) return <LoadingScreen />;
  if (passwordRecovery) return <UpdatePasswordScreen onDone={() => setPasswordRecovery(false)} />;
  if (!session) return <AuthScreen />;
  if (profilesLoading || !user) return <LoadingScreen text="Setting up your account…" />;
  if (!user.active) return <LoadingScreen text="This account has been deactivated. Contact your manager." />;
  if (costHistoryLoading || clientsLoading || feeHistoryLoading || extraFeesLoading || entriesLoading || lockedLoading || leaveLoading || rentalPropsLoading || rentalExpLoading || view === null) return <LoadingScreen />;

  return (
    <div style={{ display: "flex", minHeight: "100vh", background: C.bg, fontFamily: sans }}>
      <FontImport />
      <Sidebar user={user} view={view} setView={setView} />
      <div style={{ flex: 1, padding: "28px 32px", minWidth: 0 }}>
        {view === "timesheet" && (
          <EmployeeTimesheet user={user} clients={clients} entries={entries} refetchEntries={refetchEntries} lockedWeeks={lockedWeeks} refetchLockedWeeks={refetchLockedWeeks} />
        )}
        {view === "timeoff" && (
          <TimeOff user={user} employees={employees} leaveRequests={leaveRequests} refetchLeaveRequests={refetchLeaveRequests} />
        )}
        {user.role === "admin" && view === "leaveapprovals" && (
          <LeaveApprovals user={user} employees={employees} leaveRequests={leaveRequests} refetchLeaveRequests={refetchLeaveRequests} />
        )}
        {view === "calendar" && (
          <TeamCalendar employees={employees} leaveRequests={leaveRequests} />
        )}
        {user.role === "admin" && view === "dashboard" && (
          <AdminDashboard employees={employees} clients={clients} entries={entries} rentalProperties={rentalProperties} rentalExpenses={rentalExpenses} />
        )}
        {user.role === "admin" && view === "timesheets" && (
          <AdminTimesheets employees={employees} clients={clients} entries={entries} refetchEntries={refetchEntries} lockedWeeks={lockedWeeks} refetchLockedWeeks={refetchLockedWeeks} />
        )}
        {user.role === "admin" && view === "clients" && (
          <AdminClients clients={clients} refetchClients={refetchClients} entries={entries} />
        )}
        {user.role === "admin" && view === "employees" && (
          <AdminEmployees employees={employees} refetchEmployees={refetchEmployees} entries={entries} />
        )}
        {user.role === "admin" && view === "rental" && (
          <AdminRental properties={rentalProperties} expenses={rentalExpenses} refetchRental={refetchRental} />
        )}
        {user.role === "admin" && view === "reports" && (
          <AdminReportBuilder employees={employees} clients={clients} entries={entries} rentalExpenses={rentalExpenses} />
        )}
      </div>
    </div>
  );
}
