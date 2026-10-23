import React, { useState, useEffect, useCallback } from "react";
import {
  LayoutDashboard, Users, Building2, Clock, Plus, Pencil, Trash2,
  ChevronLeft, ChevronRight, LogOut, X, Lock, Unlock, AlertTriangle,
  Check, Search, CalendarDays, CalendarRange, ClipboardCheck, Euro, Home, SlidersHorizontal, Receipt, Wallet,
  ListChecks, MessageSquare, ArrowLeft, Phone, Mail, UserRound
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

// ---- myDATA codification (ΑΑΔΕ) used by the invoicing export for Hyper ----
// Invoice types for service fees, with the income-classification type that
// normally goes with them. Category is always category1_3 (service income).
// Confirm the combination against the AADE classification table if needed.
const MYDATA_INVOICE_TYPES = {
  "2.1": { label: "Τιμολόγιο Παροχής", incomeType: "E3_561_001", incomeTypeLabel: "Πωλήσεις αγαθών και υπηρεσιών Χονδρικές - Επιτηδευματιών" },
  "2.2": { label: "Τιμολόγιο Παροχής / Ενδοκοινοτική Παροχή Υπηρεσιών", incomeType: "E3_561_005", incomeTypeLabel: "Πωλήσεις αγαθών και υπηρεσιών Εξωτερικού Ενδοκοινοτικές" },
  "2.3": { label: "Τιμολόγιο Παροχής / Παροχή Υπηρεσιών σε λήπτη Τρίτης Χώρας", incomeType: "E3_561_006", incomeTypeLabel: "Πωλήσεις αγαθών και υπηρεσιών Εξωτερικού Τρίτες Χώρες" },
};
const MYDATA_INCOME_CATEGORY = { code: "category1_3", label: "Έσοδα από Παροχή Υπηρεσιών" };
// vatCategory 1 = 24%, 7 = 0% (needs a vatExemptionCategory)
const MYDATA_VAT_CATEGORY = { 24: 1, 0: 7 };
const MYDATA_VAT_EXEMPTIONS = [
  [1, "Χωρίς ΦΠΑ - άρθρο 2 και 3 του Κώδικα ΦΠΑ"], [2, "Χωρίς ΦΠΑ - άρθρο 5 του Κώδικα ΦΠΑ"], [3, "Χωρίς ΦΠΑ - άρθρο 17 του Κώδικα ΦΠΑ"], [4, "Χωρίς ΦΠΑ - άρθρο 18 του Κώδικα ΦΠΑ"],
  [5, "Χωρίς ΦΠΑ - άρθρο 21 του Κώδικα ΦΠΑ"], [6, "Χωρίς ΦΠΑ - άρθρο 24 του Κώδικα ΦΠΑ"], [7, "Χωρίς ΦΠΑ - άρθρο 27 του Κώδικα ΦΠΑ"], [8, "Χωρίς ΦΠΑ - άρθρο 29 του Κώδικα ΦΠΑ"],
  [9, "Χωρίς ΦΠΑ - άρθρο 30 του Κώδικα ΦΠΑ"], [10, "Χωρίς ΦΠΑ - άρθρο 31 του Κώδικα ΦΠΑ"], [11, "Χωρίς ΦΠΑ - άρθρο 32 του Κώδικα ΦΠΑ"], [14, "Χωρίς ΦΠΑ - άρθρο 33 του Κώδικα ΦΠΑ"],
  [15, "Χωρίς ΦΠΑ - άρθρο 44 του Κώδικα ΦΠΑ"], [16, "Χωρίς ΦΠΑ - άρθρο 45 του Κώδικα ΦΠΑ"], [17, "Χωρίς ΦΠΑ - άρθρο 47 του Κώδικα ΦΠΑ"], [18, "Χωρίς ΦΠΑ - άρθρο 48 του Κώδικα ΦΠΑ"],
  [19, "Χωρίς ΦΠΑ - άρθρο 54 του Κώδικα ΦΠΑ"], [24, "Χωρίς ΦΠΑ - άρθρο 8 του Κώδικα ΦΠΑ"], [25, "Χωρίς ΦΠΑ - ΠΟΛ.1029/1995"], [26, "Χωρίς ΦΠΑ - ΠΟΛ.1167/2015"], [27, "Λοιπές Εξαιρέσεις ΦΠΑ"],
];
// Service wording on the invoice line.
const INVOICE_SERVICES = { accounting: "Λογιστικές υπηρεσίες", payroll: "Υπηρεσίες μισθοδοσίας", tax: "Φορολογικές υπηρεσίες", formation: "Υπηρεσίες σύστασης εταιρείας", advisory: "Συμβουλευτικές υπηρεσίες", other: "Λοιπές υπηρεσίες" };
const EXTRA_SERVICE_BY_CC = { accounting: "accounting", tax: "tax", payroll: "payroll", other: "other", director: "advisory", finance: "advisory", rental: "other" };
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
// Other operating expenses (OPEX): insurance, accounting software, bank
// charges, equipment depreciation, AI costs, business development, other.
// Same shape as rental expenses (recurring from–to, or extra for one month;
// extras may be negative). Each one is allocated to a department, or to
// "general" (company-wide, not attributed to a department).
const OPEX_CATEGORIES = ["accounting_software", "other_software", "database", "ai_costs", "insurance", "car_insurance", "car_depreciation", "depreciation", "website", "emails", "server_cost", "business_development", "bank_costs", "efka_director", "health_director", "pos_charges", "other"];
const OPEX_LABELS = {
  accounting_software: "Accounting software (λογιστικά προγράμματα)", other_software: "Other software (PDF, Excel κ.λπ.)", database: "Database", ai_costs: "AI costs",
  insurance: "Insurance — office (ασφάλεια γραφείου)", car_insurance: "Insurance — car (ασφάλεια αυτοκινήτου)",
  car_depreciation: "Depreciation — car (αποσβέσεις αυτοκινήτου)", depreciation: "Depreciation — office (αποσβέσεις γραφείου)",
  website: "Website", emails: "Emails", server_cost: "Server costs", business_development: "Business development", bank_costs: "Bank charges (τραπεζικά κόστη)", efka_director: "EFKA — director (ΕΦΚΑ διαχειριστή)", health_director: "Health insurance — director (ασφάλεια υγείας διαχειριστή)", pos_charges: "POS charges (προμήθειες POS)", other: "Other (note required)",
};
const OPEX_GENERAL = "general";
OPEX_LABELS.rental = "Rental (ενοίκια, utilities, κοινόχρηστα)";
const OPEX_DISPLAY_CATEGORIES = ["rental", ...OPEX_CATEGORIES]; // "rental" comes from the Rental page
const opexDepartmentLabel = (key) => key === OPEX_GENERAL ? "General (company-wide)" : (DEPARTMENTS.find((d) => d.key === key)?.label || key);
const mapOpex = (r) => ({ id: r.id, kind: r.kind === "recurring" ? "recurring" : "extra", month: r.month || null, fromMonth: r.from_month || null, toMonth: r.to_month || null, category: r.category, department: r.department || OPEX_GENERAL, amount: Number(r.amount), note: r.note || "" });
// Operating expenses accrued over a date range (daily accrual by month, like rental expenses).
// Rental expenses (rent, utilities, common expenses) count here too — they are
// operating expenses, attributed to the Rental department.
function opexForRange(expenses, rangeStart, rangeEnd, rentalExpenses = []) {
  const byCategory = Object.fromEntries(OPEX_DISPLAY_CATEGORIES.map((k) => [k, 0]));
  const byDepartment = { accounting_team: 0, management_team: 0, rental: 0, [OPEX_GENERAL]: 0 };
  let total = 0;
  let d = new Date(rangeStart);
  while (d <= rangeEnd) {
    const monthKey = monthKeyOf(d);
    const days = endOfMonth(d).getDate();
    (expenses || []).filter((x) => rentalExpenseAppliesTo(x, monthKey)).forEach((x) => {
      const amt = (Number(x.amount) || 0) / days;
      byCategory[OPEX_CATEGORIES.includes(x.category) ? x.category : "other"] += amt;
      const dep = x.department in byDepartment ? x.department : OPEX_GENERAL;
      byDepartment[dep] += amt;
      total += amt;
    });
    d = addDays(d, 1);
  }
  const rent = rentalExpensesForRange(rentalExpenses, rangeStart, rangeEnd).total;
  byCategory.rental += rent; byDepartment.rental += rent; total += rent;
  return { total, byCategory, byDepartment };
}
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
              <div style={{ fontFamily: sans, fontSize: 13, color: C.inkMuted }}>You can continue into GLORIA now.</div>
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
            {headers.map((h, hi) => (
              <th key={hi} style={{
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
    { key: "mytasks", label: "My Tasks", icon: ListChecks },
    { key: "timesheet", label: "My Timesheet", icon: Clock },
    { key: "timeoff", label: "Time Off", icon: CalendarDays },
    { key: "calendar", label: "Calendar", icon: CalendarRange },
    { key: "clients", label: "Clients", icon: Building2 },
    { key: "billing", label: "Billing", icon: Receipt },
  ];
  const adminNav = [
    { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { key: "timesheet", label: "My Timesheet", icon: Clock },
    { key: "timesheets", label: "All Timesheets", icon: Clock },
    { key: "timeoff", label: "Time Off", icon: CalendarDays },
    { key: "leaveapprovals", label: "Leave Approvals", icon: ClipboardCheck },
    { key: "calendar", label: "Calendar", icon: CalendarRange },
    { key: "clients", label: "Clients", icon: Building2 },
    { key: "tasks", label: "Tasks", icon: ListChecks },
    { key: "employees", label: "Employees", icon: Users },
    { key: "rental", label: "Rental", icon: Home },
    { key: "opex", label: "Operating Expenses", icon: Wallet },
    { key: "reports", label: "Reports", icon: SlidersHorizontal },
    { key: "billing", label: "Billing", icon: Receipt },
    { key: "compliance", label: "AML Compliance", icon: AlertTriangle },
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
          const active = view === it.key || (view === "clientcard" && it.key === "clients");
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
const STABLE_ORDER = { locked_weeks: ["employee_id", "week_start"], app_settings: ["key"], client_aml: ["client_id"] };
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
  // Rental expenses are operating expenses now (below gross profit), so cost here is labor only.
  const cost = laborCost;
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
const mapClientRow = (r) => ({
  id: r.id, name: r.name, fixedFee: r.fixed_fee === null ? null : Number(r.fixed_fee), endDate: r.end_date || null, active: r.active, createdAt: r.created_at ? r.created_at.slice(0, 10) : null,
  afm: r.afm || "", doy: r.doy || "", legalForm: r.legal_form || "", address: r.address || "", email: r.email || "", phone: r.phone || "",
  accountManagerId: r.account_manager_id || null, generalNotes: r.general_notes || "",
  representativeName: r.representative_name || "", representativeAfm: r.representative_afm || "", representativeIdType: r.representative_id_type || "", representativeIdNumber: r.representative_id_number || "", representativeEmail: r.representative_email || "", representativePhone: r.representative_phone || "",
  lifecycle: r.lifecycle || "live", goLiveDate: r.go_live_date || null, engagementDate: r.engagement_date || null, leadSource: r.lead_source || "", expectedFee: r.expected_fee === null || r.expected_fee === undefined ? null : Number(r.expected_fee),
  contractStatus: r.contract_status || "none", contractSignedOn: r.contract_signed_on || null, contractVersion: Number(r.contract_version) || 1, contractSpecialTerms: r.contract_special_terms || "", contractPaymentTerms: r.contract_payment_terms || "",
  contractServicesText: r.contract_services_text || "", contractPayrollBaseFee: r.contract_payroll_base_fee === null || r.contract_payroll_base_fee === undefined ? null : Number(r.contract_payroll_base_fee),
  contractPayrollIncluded: r.contract_payroll_included === null || r.contract_payroll_included === undefined ? null : Number(r.contract_payroll_included), contractPayrollExtraFee: r.contract_payroll_extra_fee === null || r.contract_payroll_extra_fee === undefined ? null : Number(r.contract_payroll_extra_fee),
  contractHourlyRate: r.contract_hourly_rate === null || r.contract_hourly_rate === undefined ? null : Number(r.contract_hourly_rate), contractNoticeMonths: r.contract_notice_months === null || r.contract_notice_months === undefined ? null : Number(r.contract_notice_months),
  contractIncludeAml: r.contract_include_aml !== false, contractStartDate: r.contract_start_date || null, contractOtherFeesText: r.contract_other_fees_text || "",
  vatRate: r.vat_rate === null || r.vat_rate === undefined ? 24 : Number(r.vat_rate),
  invoiceType: r.invoice_type && MYDATA_INVOICE_TYPES[r.invoice_type] ? r.invoice_type : "2.1", vatExemptionCategory: r.vat_exemption_category === null || r.vat_exemption_category === undefined ? null : Number(r.vat_exemption_category), fixedService: r.fixed_service && INVOICE_SERVICES[r.fixed_service] ? r.fixed_service : "accounting",
  annualFeeEnabled: r.annual_fee_enabled === true, annualFee: r.annual_fee === null || r.annual_fee === undefined ? null : Number(r.annual_fee), annualFeeMonth: r.annual_fee_month || "", annualFeeCc: r.annual_fee_cc || "accounting", annualFeeExtraId: r.annual_fee_extra_id || null,
  formationFeeEnabled: r.formation_fee_enabled === true, formationFee: r.formation_fee === null || r.formation_fee === undefined ? null : Number(r.formation_fee), formationFeeMonth: r.formation_fee_month || "", formationFeeCc: r.formation_fee_cc || "other", formationFeeExtraId: r.formation_fee_extra_id || null,
});
const mapDocument = (r) => ({ id: r.id, clientId: r.client_id, category: r.category || "other", fileName: r.file_name, filePath: r.file_path, fileSize: Number(r.file_size) || 0, mimeType: r.mime_type || "", note: r.note || "", uploadedBy: r.uploaded_by || null, uploadedAt: r.uploaded_at });
const DOCUMENT_BUCKET = "client-documents";
const CONTRACT_STATUS_LABELS = { none: "No agreement", draft: "Draft", sent: "Sent for signature", signed: "Signed" };
const CONTRACT_STATUS_TONE = { none: "neutral", draft: "warn", sent: "warn", signed: "accent" };
const DEFAULT_FIRM = { name: "ACCOUNTING ACCO E.E.", legalFormText: "Ετερόρρυθμη Εταιρεία", afm: "802815063", doy: "ΚΕΦΟΔΕ ΑΤΤΙΚΗΣ", address: "Κάνιγγος 2, ΤΚ 10677, Αθήνα", representative: "Χάλαρης Σπυρίδων Δαυίδ", representativeId: "ΑΗ505813", representativeAfm: "141019020", email: "", phone: "", gemi: "" };
const CONTRACT_DEFAULT_SERVICES = `Τήρηση λογιστικών βιβλίων.
Υποβολή μηνιαίων/τριμηνιαίων φορολογικών δηλώσεων (ΦΠΑ, Παρακρατούμενων, Συμφωνητικά, Ψηφιακά Τέλη κλπ).
Έλεγχος συμφωνίας myDATA με τα βιβλία της επιχείρησης.
Υποβολή ετήσιας δήλωσης φορολογίας εισοδήματος (έντυπα Ν και Ε3).
Λοιπές δηλώσεις φορολογικής συμμόρφωσης στην ΑΑΔΕ.
Υποβολή λοιπών αιτήσεων στο ΓΕΜΗ και γενικότερη επικοινωνία με το επιμελητήριο.`;
const CONTRACT_DEFAULT_PAYMENT = "Η τιμολόγηση πραγματοποιείται μηνιαίως 10 ημέρες πριν από τη λήξη του μήνα και η πληρωμή από τον Εντολέα θα γίνεται μέχρι την τελευταία εργάσιμη του επόμενου μήνα.";
const GREEK_MONTHS_GEN = ["Ιανουαρίου", "Φεβρουαρίου", "Μαρτίου", "Απριλίου", "Μαΐου", "Ιουνίου", "Ιουλίου", "Αυγούστου", "Σεπτεμβρίου", "Οκτωβρίου", "Νοεμβρίου", "Δεκεμβρίου"];
const fmtGreekDate = (key) => { if (!key) return "____/____/______"; const d = fromKey(key); return `${d.getDate()}${d.getDate() === 1 ? "η" : "η"} ${GREEK_MONTHS_GEN[d.getMonth()]} ${d.getFullYear()}`; };
const fmtEurGr = (n) => Number(n).toLocaleString("el-GR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €";
const fmtBytes = (n) => n > 1048576 ? (n / 1048576).toFixed(1) + " MB" : n > 1024 ? Math.round(n / 1024) + " KB" : n + " B";
const escapeHtml = (v) => String(v ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
// Opens a print window with the given HTML so the user can "Save as PDF".
// No extra library needed; every browser prints to PDF.
function printHtml(title, bodyHtml, { draft = false, editable = false } = {}) {
  const w = window.open("", "_blank");
  if (!w) { window.alert("Allow pop-ups for this site to print / save as PDF."); return; }
  w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title>
  <style>
    body { font-family: Georgia, "Times New Roman", serif; font-size: 11.5pt; color: #111; margin: 36px 44px; line-height: 1.45; }
    h1 { font-size: 17pt; margin: 0 0 4px; } h2 { font-size: 13pt; margin: 22px 0 6px; border-bottom: 1px solid #999; padding-bottom: 3px; }
    h3 { font-size: 11.5pt; margin: 14px 0 4px; } p { margin: 6px 0; } .muted { color: #555; font-size: 10pt; }
    table { border-collapse: collapse; width: 100%; margin: 6px 0 10px; font-size: 10.5pt; } th, td { border: 1px solid #bbb; padding: 4px 6px; vertical-align: top; text-align: left; }
    th { background: #f0f0f0; } .kv td:first-child { width: 34%; color: #444; } .sig { display: flex; gap: 40px; margin-top: 40px; } .sig div { flex: 1; border-top: 1px solid #333; padding-top: 6px; font-size: 10pt; }
    .stamp { position: fixed; top: 40%; left: 10%; font-size: 64pt; color: rgba(200,0,0,0.12); transform: rotate(-25deg); font-weight: bold; pointer-events: none; }
    .footer { margin-top: 28px; font-size: 9pt; color: #666; border-top: 1px solid #ccc; padding-top: 6px; }
    @media print { .noprint { display: none; } }
  </style></head><body>
  ${draft ? '<div class="stamp">DRAFT</div>' : ""}
  <div class="noprint" style="margin-bottom:14px;font-family:sans-serif;font-size:12px;color:#555">${editable ? "<strong>Μπορείτε να πληκτρολογήσετε απευθείας πάνω στο έντυπο και να πατήσετε τα ☐ για επιλογή.</strong> " : ""}Use your browser's Print and choose "Save as PDF". <button onclick="window.print()">Print / Save as PDF</button></div>
  <div${editable ? ' contenteditable="true" style="outline:none"' : ""}>${bodyHtml}</div>
  <script>document.addEventListener("click", function (ev) { var t = ev.target; if (t && t.classList && t.classList.contains("cb")) { t.textContent = t.textContent.trim() === "☒" ? "☐" : "☒"; } });</script>
  </body></html>`);
  w.document.close();
  if (!editable) setTimeout(() => { try { w.focus(); w.print(); } catch (e) { /* user can press the button */ } }, 400);
}
const mapTemplate = (r) => ({ id: r.id, stage: r.stage || "setup", sortOrder: Number(r.sort_order) || 0, title: r.title, description: r.description || "", dueDays: Number(r.due_days) || 0, active: r.active !== false, appliesTo: r.applies_to || "all" });
const mapAml = (r) => ({
  clientId: r.client_id, clientType: r.client_type || "legal", gemiNumber: r.gemi_number || "", incorporationDate: r.incorporation_date || "",
  representativeName: r.representative_name || "", representativeIdType: r.representative_id_type || "", representativeIdNumber: r.representative_id_number || "", representativeAfm: r.representative_afm || "",
  activity: r.activity || "", countries: r.countries || "", relationshipPurpose: r.relationship_purpose || "", services: Array.isArray(r.services) ? r.services : [],
  sourceOfFunds: r.source_of_funds || "", pepStatus: r.pep_status || "none", pepDetails: r.pep_details || "", sanctionsCheckedOn: r.sanctions_checked_on || "", sanctionsResult: r.sanctions_result || "",
  riskFactors: Array.isArray(r.risk_factors) ? r.risk_factors : [], riskScore: Number(r.risk_score) || 0, riskLevel: r.risk_level || "low", riskOverride: r.risk_override || "", riskOverrideReason: r.risk_override_reason || "",
  ddLevel: r.dd_level || "standard", documents: r.documents && typeof r.documents === "object" ? r.documents : {}, status: r.status || "not_started",
  completedOn: r.completed_on || "", completedBy: r.completed_by || null, approvedBy: r.approved_by || null, approvedOn: r.approved_on || "", nextReviewOn: r.next_review_on || "", notes: r.notes || "",
  entityType: r.entity_type || "", infoSources: r.info_sources || "", proxyName: r.proxy_name || "", proxyIdNumber: r.proxy_id_number || "", decision: r.decision || "", decisionReason: r.decision_reason || "", uboActionsRecord: r.ubo_actions_record || "",
  riskFactorNotes: r.risk_factor_notes && typeof r.risk_factor_notes === "object" ? r.risk_factor_notes : {}, assessmentType: r.assessment_type || "initial", assessmentDate: r.assessment_date || "", sheetVersion: Number(r.sheet_version) || 1, riskScoreB: Number(r.risk_score_b) || 0,
  uboChangeDate: r.ubo_change_date || "", uboFiledDate: r.ubo_filed_date || "",
});
const mapUbo = (r) => ({ id: r.id, clientId: r.client_id, name: r.name, afm: r.afm || "", idNumber: r.id_number || "", nationality: r.nationality || "", residenceCountry: r.residence_country || "", ownershipPct: r.ownership_pct === null || r.ownership_pct === undefined ? null : Number(r.ownership_pct), controlType: r.control_type || "", isPep: !!r.is_pep, verifiedOn: r.verified_on || "", verificationDoc: r.verification_doc || "", note: r.note || "" });

// Onboarding stages (order matters for the checklist display).
const ONBOARDING_STAGES = [
  { key: "aml", label: "Due diligence (ν. 4557/2018)" },
  { key: "contract", label: "Contract & authorisations" },
  { key: "formation", label: "Company formation (new entity)" },
  { key: "setup", label: "Client setup" },
];
const APPLIES_TO_LABELS = { all: "All clients", capital: "Α.Ε. / Ε.Π.Ε. / Ι.Κ.Ε. only", if_applicable: "If applicable" };
const isCapitalCompany = (legalForm) => ["Α.Ε.", "Ε.Π.Ε.", "Ι.Κ.Ε."].includes(legalForm);
const stageLabel = (k) => ONBOARDING_STAGES.find((s) => s.key === k)?.label || k;

// AML risk model — "Μοντέλο Α" of the Taxheaven guide (Κουλογιάννης, 8/2026):
// four axes mirroring παρ. 1 άρθρου 35 ν. 4557/2018. Annex I factors subtract
// points, Annex II factors add; a few cases override everything and force
// high risk / enhanced due diligence regardless of the sum. The points are
// a firm-level methodology choice, not a statutory scale.
const AML_AXES = [
  { key: "A", label: "Axis A — Type and nature of the client" },
  { key: "B", label: "Axis B — Geography (worst of seat, shareholders, UBO, counterparties)" },
  { key: "C", label: "Axis C — Services and products" },
  { key: "D", label: "Axis D — Transactions and delivery channel" },
];
const AML_RISK_FACTORS = [
  // Axis A
  { key: "a_listed", axis: "A", points: -2, ref: "I-1α", label: "Listed on an EU regulated market (or equivalent)" },
  { key: "a_public", axis: "A", points: -2, ref: "I-1β", label: "Public authority, ΝΠΔΔ, majority state-owned, EU body" },
  { key: "a_unusual", axis: "A", points: 3, ref: "II-1α", label: "Relationship developing in unusual circumstances" },
  { key: "a_asset_vehicle", axis: "A", points: 3, ref: "II-1γ", label: "Personal asset-holding vehicle" },
  { key: "a_nominee", axis: "A", points: 4, ref: "II-1δ", label: "Nominee shareholders or bearer shares" },
  { key: "a_cash", axis: "A", points: 3, ref: "II-1ε", label: "Cash-intensive business" },
  { key: "a_complex", axis: "A", points: 4, ref: "II-1στ", label: "Unusual or excessively complex ownership structure" },
  { key: "a_golden_visa", axis: "A", points: 4, ref: "II-1ζ", label: "Third-country national applying for residence/citizenship by investment" },
  { key: "a_ubo_unknown", axis: "A", points: 4, ref: "άρ. 3 παρ. 17", label: "Beneficial owner not identified despite exhausting all means" },
  { key: "a_pep", axis: "A", override: true, ref: "άρ. 18", label: "Client or UBO is a PEP, close relative or close associate" },
  // Axis B
  { key: "b_eu", axis: "B", points: -1, ref: "I-3α", label: "Greece or another EU member state" },
  { key: "b_fatf_ok", axis: "B", points: 0, ref: "I-3β/γ", label: "Third country with documented positive FATF assessment" },
  { key: "b_third", axis: "B", points: 2, ref: "—", label: "Third country without findings but without positive assessment" },
  { key: "b_corruption", axis: "B", points: 3, ref: "II-3β", label: "Country with high corruption or organised crime (credible sources)" },
  { key: "b_weak_aml", axis: "B", points: 4, ref: "II-3α", label: "Country with established lack of effective AML/CFT systems" },
  { key: "b_eu_list", axis: "B", override: true, ref: "άρ. 16, 16Α", label: "High-risk third country on the European Commission list" },
  { key: "b_sanctions", axis: "B", override: true, ref: "II-3γ", label: "Country under EU or UN sanctions / restrictive measures" },
  { key: "b_terror", axis: "B", override: true, ref: "II-3δ", label: "Country financing or supporting terrorist activities" },
  // Axis C
  { key: "c_simple", axis: "C", points: -1, ref: "I-2ε", label: "Bookkeeping and filings only, small scale, full transparency of identity" },
  { key: "c_formation", axis: "C", points: 2, ref: "—", label: "Company formation, corporate changes or registered-office services" },
  { key: "c_private_banking", axis: "C", points: 2, ref: "II-2α", label: "Client maintains a private banking relationship" },
  { key: "c_crypto", axis: "C", points: 3, ref: "II-2ε", label: "New products or technologies (e.g. crypto-assets)" },
  { key: "c_sensitive_goods", axis: "C", points: 3, ref: "II-2στ", label: "Oil, precious metals, tobacco, cultural goods, ivory, protected species" },
  { key: "c_anonymity", axis: "C", points: 4, ref: "II-2β", label: "Products or transactions favouring anonymity" },
  // Axis D
  { key: "d_remote_eid", axis: "D", points: 0, ref: "II-2γ (neutralised)", label: "Remote relationship with recognised e-ID or equivalent safeguards" },
  { key: "d_proxy", axis: "D", points: 1, ref: "—", label: "Client acts through an authorised person / proxy" },
  { key: "d_remote", axis: "D", points: 3, ref: "II-2γ", label: "Remote relationship without eIDAS-level safeguards" },
  { key: "d_third_party_pay", axis: "D", points: 4, ref: "II-2δ", label: "Payments from unrelated third parties" },
  { key: "d_complex_tx", axis: "D", points: 5, ref: "άρ. 16 παρ. 3", label: "Complex, unusually large or unusual-pattern transactions, no apparent purpose" },
];
// Scale (Β.5.2.2 of the guide): ≤0 with at least one Annex I factor → low;
// +1..+5, or ≤0 without any Annex I factor → medium; ≥+6 or any override → high.
function amlAssess(factorKeys) {
  const chosen = AML_RISK_FACTORS.filter((f) => factorKeys.includes(f.key));
  const score = chosen.reduce((s, f) => s + (f.points || 0), 0);
  const forced = chosen.some((f) => f.override);
  const hasAnnexI = chosen.some((f) => (f.points || 0) < 0);
  let level;
  if (forced || score >= 6) level = "high";
  else if (score >= 1 || !hasAnnexI) level = "medium";
  else level = "low";
  return { score, level, forced, hasAnnexI, overrides: chosen.filter((f) => f.override) };
}
const AML_DD_FOR_LEVEL = { low: "simplified", medium: "standard", high: "enhanced" };
// Review cadence from the ΑΑΔΕ seminar material quoted in the guide: at
// least every 2 years for low risk, at least yearly otherwise (and on every
// change for high risk).
const AML_REVIEW_MONTHS = { low: 24, medium: 12, high: 12 };
const AML_LEVEL_LABELS = { low: "Low", medium: "Medium", high: "High" };
const AML_LEVEL_TONE = { low: "accent", medium: "warn", high: "danger" };
const AML_DD_LABELS = { simplified: "Simplified (άρ. 15)", standard: "Standard (άρ. 13)", enhanced: "Enhanced (άρ. 16–18)" };
const AML_STATUS_LABELS = { not_started: "Not started", in_progress: "In progress", completed: "Completed", declined: "Declined" };
const AML_STATUS_TONE = { not_started: "neutral", in_progress: "warn", completed: "accent", declined: "danger" };
// Identification documents per ΠΟΛ.1200/2018 tables A–D, plus the file
// items the guide's new-client flow requires (KYC form with GDPR notice,
// UBO register proof, engagement letter).
const AML_ENTITY_TYPES = { natural: "Natural person (Table A)", capital: "Capital company — Α.Ε., Ε.Π.Ε., Ι.Κ.Ε. (Table B)", partnership: "Partnership — Ο.Ε., Ε.Ε. (Table C)", other: "Other legal person / entity (Table D)" };
const amlEntityTypeFor = (legalForm) => ["Ατομική", "Ιδιώτης"].includes(legalForm) ? "natural" : ["Α.Ε.", "Ε.Π.Ε.", "Ι.Κ.Ε."].includes(legalForm) ? "capital" : ["Ο.Ε.", "Ε.Ε."].includes(legalForm) ? "partnership" : "other";
const AML_DOCUMENTS = [
  { key: "kyc_form", label: "Signed KYC questionnaire incl. GDPR notice (άρ. 31 παρ. 2)", types: ["natural", "capital", "partnership", "other"] },
  { key: "rep_id", label: "ID / passport / residence permit of client or legal representative (Table A)", types: ["natural", "capital", "partnership", "other"] },
  { key: "address", label: "Proof of address — recent utility bill or lease (Table A)", types: ["natural"] },
  { key: "occupation", label: "Occupation evidence — employer certificate, payslip, professional ID, insurance certificate (Table A)", types: ["natural"] },
  { key: "afm_natural", label: "ΑΦΜ certificate (βεβαίωση απόδοσης ΑΦΜ)", types: ["natural"] },
  { key: "articles", label: "Articles of association published in ΓΕΜΗ, with amendments", types: ["capital", "partnership"] },
  { key: "legalisation", label: "Legalisation documents with amendments", types: ["other"] },
  { key: "appointment", label: "Appointment of managers / legal representative (ΓΕΜΗ)", types: ["capital"] },
  { key: "afm_entity", label: "ΑΑΔΕ start-of-business certificate (βεβαίωση έναρξης μη φυσικού προσώπου)", types: ["capital", "partnership", "other"] },
  { key: "shareholders", label: "Shareholder / partner register (ownership mapping)", types: ["capital", "partnership", "other"] },
  { key: "ubo_registry", label: "Proof of registration in the Central UBO Register (άρ. 14 παρ. 1)", types: ["capital", "partnership", "other"] },
  { key: "ubo_ids", label: "Table A documents for every beneficial owner", types: ["capital", "partnership", "other"] },
  { key: "proxy", label: "Authorised person: ID and authorisation (if acting via proxy)", types: ["natural", "capital", "partnership", "other"], optional: true },
  { key: "source_of_funds", label: "Source of wealth / origin of funds evidence (enhanced due diligence)", types: ["natural", "capital", "partnership", "other"], enhancedOnly: true },
  { key: "engagement", label: "Signed engagement letter describing the services", types: ["natural", "capital", "partnership", "other"] },
];
const AML_SERVICES = ["Bookkeeping", "Tax filings", "Payroll", "Advisory", "Representation before authorities", "Company formation / corporate changes", "Registered office"];
// ---- Model B (0–100): same factors, per-axis points, highest + 5 per extra
// factor, capped at 25 per axis. Annex I factors keep the axis at 0.
const AML_MODEL_B_POINTS = {
  a_listed: 0, a_public: 0, a_unusual: 15, a_asset_vehicle: 15, a_cash: 15, a_nominee: 20, a_complex: 20, a_golden_visa: 20, a_ubo_unknown: 25,
  b_eu: 0, b_fatf_ok: 5, b_third: 10, b_corruption: 15, b_weak_aml: 20,
  c_simple: 0, c_formation: 10, c_private_banking: 10, c_crypto: 15, c_sensitive_goods: 15, c_anonymity: 20,
  d_remote_eid: 0, d_proxy: 5, d_remote: 15, d_third_party_pay: 20, d_complex_tx: 20,
};
const AML_MODEL_B_BASE = { A: 5, C: 5 }; // "βάση" rows when nothing else is ticked on the axis
function amlAssessB(factorKeys) {
  const chosen = AML_RISK_FACTORS.filter((f) => factorKeys.includes(f.key) && !f.override);
  const forced = AML_RISK_FACTORS.some((f) => f.override && factorKeys.includes(f.key));
  const hasAnnexI = chosen.some((f) => (f.points || 0) < 0);
  const hasAnnexII = chosen.some((f) => (f.points || 0) > 0 && !["c_formation", "b_third", "d_proxy"].includes(f.key));
  const axes = {};
  AML_AXES.forEach((ax) => {
    const pts = chosen.filter((f) => f.axis === ax.key).map((f) => AML_MODEL_B_POINTS[f.key] ?? 0).sort((x, y) => y - x);
    let v = pts.length ? pts[0] + 5 * (pts.length - 1) : (AML_MODEL_B_BASE[ax.key] || 0);
    if (pts.length && pts[0] === 0 && chosen.some((f) => f.axis === ax.key && (f.points || 0) < 0)) v = 0;
    axes[ax.key] = Math.min(25, v);
  });
  const score = Object.values(axes).reduce((s, v) => s + v, 0);
  let level;
  if (forced || score >= 46) level = "high";
  else if (score >= 16 || !hasAnnexI || hasAnnexII) level = "medium";
  else level = "low";
  return { score, axes, level, forced, hasAnnexI, hasAnnexII };
}
const AML_ASSESSMENT_TYPES = { initial: "Αρχική", annual: "Ετήσια", extraordinary: "Έκτακτη" };

// ---- The 23 ΑΑΔΕ audit questions (self-assessment), with the guide's pointer.
const AML_AUDIT_QUESTIONS = [
  { n: 1, group: "Εσωτερικές διαδικασίες / πολιτικές", q: "Υπάρχει συστηματική ενημέρωση/εκπαίδευση του προσωπικού για θέματα ξεπλύματος; (αρ. 37)", fix: "Βήμα 12 — εκπαίδευση και αρχείο εκπαίδευσης" },
  { n: 2, group: "Εσωτερικές διαδικασίες / πολιτικές", q: "Υπάρχει γραπτή επικαιροποιημένη πολιτική (εγχειρίδιο πολιτικών και διαδικασιών); (αρ. 35)", fix: "Βήμα 10 — εγχειρίδιο με ημερομηνία έγκρισης" },
  { n: 3, group: "Εσωτερικές διαδικασίες / πολιτικές", q: "Έχει οριστεί υπεύθυνος συμμόρφωσης; (αρ. 35)", fix: "Βήμα 1 — γραπτή πράξη ορισμού" },
  { n: 4, group: "Εσωτερικές διαδικασίες / πολιτικές", q: "Υπάρχουν διαδικασίες εσωτερικής καταγγελίας παραβάσεων μέσω ειδικού, ανεξάρτητου και ανώνυμου διαύλου; (αρ. 35 παρ. 3β)", fix: "Βήμα 11 — δίαυλος ή γραπτή τεκμηρίωση απουσίας προσωπικού" },
  { n: 5, group: "Εσωτερικές διαδικασίες / πολιτικές", q: "Αν ανήκετε σε όμιλο, εφαρμόζονται πολιτικές σε επίπεδο ομίλου; (αρ. 36)", fix: "Αν δεν ανήκετε, απαντάτε «δεν συντρέχει» — όχι κενό", na: true },
  { n: 6, group: "Μέτρα δέουσας επιμέλειας", q: "Υπάρχει διαδικασία επαλήθευσης εγγράφων ταυτοποίησης (Γ.Ε.ΜΗ., ΦΕΚ σύστασης, ΠΟΛ.1200/2018); (αρ. 13 παρ. 1α)", fix: "Βήμα 4 — έντυπο KYC με τους Πίνακες Α–Δ" },
  { n: 7, group: "Μέτρα δέουσας επιμέλειας", q: "Υπάρχει διαδικασία εξακρίβωσης πραγματικών δικαιούχων; (αρ. 13 παρ. 1β)", fix: "Βήμα 5 — έντυπο ΠΔ και χαρτογράφηση δομής" },
  { n: 8, group: "Μέτρα δέουσας επιμέλειας", q: "Τηρείτε αρχείο με τα στοιχεία ταυτοποίησης των πελατών; (αρ. 30, 31)", fix: "Ατομικοί φάκελοι πελατών, έντυποι ή ψηφιακοί" },
  { n: 9, group: "Μέτρα δέουσας επιμέλειας", q: "Έχει πραγματοποιηθεί συναλλαγή με πελάτη πριν την ολοκλήρωση της ταυτοποίησης; (αρ. 14 παρ. 1 και 2)", fix: "Τεκμηριώστε ημερομηνίες ταυτοποίησης πριν από την έναρξη κάθε σχέσης", invert: true },
  { n: 10, group: "Μέτρα δέουσας επιμέλειας", q: "Υπάρχει διαδικασία για ΠΕΠ / στενούς συγγενείς / στενούς συνεργάτες και εφαρμόζονται μέτρα αυξημένης επιμέλειας; (αρ. 3 παρ. 9-10, αρ. 16, 18)", fix: "Ενότητα Β.7.3 — ερώτημα ΠΕΠ στο KYC και τα 4 μέτρα του άρθρου 18" },
  { n: 11, group: "Μέτρα δέουσας επιμέλειας", q: "Έχει απορριφθεί συνεργασία λόγω υποψίας σύνδεσης με ΞΧ/ΧΤ; (αρ. 23)", fix: "Τηρήστε αρχείο απορρίψεων και διακοπών με αιτιολογία", na: true },
  { n: 12, group: "Μέτρα δέουσας επιμέλειας", q: "Υπάρχει πολιτική επικαιροποίησης της δέουσας επιμέλειας για τους υφιστάμενους πελάτες; (αρ. 13 παρ. 7)", fix: "Βήμα 14 — πρόγραμμα επανεξέτασης με ημερομηνίες" },
  { n: 13, group: "Μέτρα δέουσας επιμέλειας", q: "Υπάρχει διαδικασία εντοπισμού ασυνήθιστων συναλλαγών; (αρ. 16)", fix: "Ενότητα Β.7.4 — γραπτή ακολουθία εντοπισμού και κλιμάκωσης" },
  { n: 14, group: "Μέτρα δέουσας επιμέλειας", q: "Τηρείτε αρχείο εγγράφων για πέντε έτη μετά τη λήξη της σχέσης; (αρ. 30 παρ. 3)", fix: "Πολιτική τήρησης 5ετίας και διαγραφής δεδομένων μετά τη λήξη" },
  { n: 15, group: "Ανάλυση κινδύνου", q: "Γίνεται εκτίμηση κινδύνου κάθε πελάτη, με μέθοδο, τεκμηρίωση και λίστα όλων των αξιολογημένων πελατών; (αρ. 13 παρ. 9, αρ. 35 παρ. 1-2)", fix: "Βήματα 7-9 — μεθοδολογία, φύλλο ανά πελάτη και συγκεντρωτική κατάσταση" },
  { n: 16, group: "Ανάλυση κινδύνου", q: "Η εκτίμηση κινδύνου καλύπτει τύπο πελάτη, γεωγραφία, τρόπους πληρωμής, προϊόντα και υπηρεσίες; (αρ. 15, 16)", fix: "Βήμα 7 — οι τέσσερις άξονες της μεθοδολογίας" },
  { n: 17, group: "Ανάλυση κινδύνου", q: "Έχουν συνεκτιμηθεί όλοι οι παράγοντες των Παραρτημάτων Ι και ΙΙ; (αρ. 15, 16)", fix: "Ενότητες Β.5.1-Β.5.2 — αντιστοίχιση με τους κωδικούς των Παραρτημάτων" },
  { n: 18, group: "Αποστολή αναφορών", q: "Υπάρχει πολιτική υποβολής αναφορών προς την Αρχή και αρχείο αναφορών; (αρ. 22)", fix: "Πολιτική STR, αρχείο αναφορών και αρχείο περιστατικών που εξετάστηκαν χωρίς αναφορά (αρ. 30 παρ. 1γ)" },
  { n: 19, group: "Αποστολή αναφορών", q: "Τηρείται η απαγόρευση γνωστοποίησης (tipping-off) και ενημερώνονται οι υπάλληλοι; (αρ. 27)", fix: "Ρήτρα στο εγχειρίδιο και γραπτή ενημέρωση προσωπικού" },
  { n: 20, group: "Λοιπές", q: "Έχετε καταχωρίσει στο Μητρώο Πραγματικών Δικαιούχων τα στοιχεία της δικής σας εταιρείας; (αρ. 20)", fix: "Βήμα 2 — ΚΜΠΔ και ειδικό μητρώο στην έδρα" },
  { n: 21, group: "Λοιπές", q: "Έχετε συμπληρώσει τον Πίνακα Ι΄ του εντύπου Ε3 σε κάθε φορολογικό έτος;", fix: "Βήμα 2 — ελέγξτε τις υποβληθείσες δηλώσεις σας" },
  { n: 22, group: "Λοιπές", q: "Με τι συχνότητα επικαιροποιείτε τις πληροφορίες της επιχειρηματικής σχέσης; (αρ. 13 παρ. 1δ)", fix: "Γραπτός πίνακας συχνοτήτων ανά κατηγορία κινδύνου" },
  { n: 23, group: "Λοιπές", q: "Συλλέγονται πληροφορίες για αντικείμενο και σκοπό της σχέσης, και από ποιες πηγές; (αρ. 13 παρ. 1γ)", fix: "Βήμα 6 — σημείωμα σκοπού και φύσης ανά πελάτη με αναφορά πηγών" },
];
// ---- The 15-step harmonisation roadmap (five stages).
const AML_STEPS = [
  { n: 1, stage: "Προετοιμασία", title: "Ορίστε Υπεύθυνο Συμμόρφωσης", out: "Γραπτή πράξη ορισμού, με ημερομηνία και στοιχεία επικοινωνίας", q: [3] },
  { n: 2, stage: "Προετοιμασία", title: "Τακτοποιήστε τις δικές σας υποχρεώσεις", out: "Καταχώριση του γραφείου στο ΚΜΠΔ, Πίνακας Ι΄ του Ε3", q: [20, 21] },
  { n: 3, stage: "Συγκέντρωση στοιχείων", title: "Απογραφή του πελατολογίου", out: "Ενιαίος κατάλογος πελατών με νομική μορφή, υπηρεσίες και προφανείς ενδείξεις", q: [] },
  { n: 4, stage: "Συγκέντρωση στοιχείων", title: "Συλλέξτε τα έγγραφα ταυτοποίησης", out: "Πίνακες Α–Δ της ΠΟΛ.1200/2018, έντυπο KYC", q: [6, 8] },
  { n: 5, stage: "Συγκέντρωση στοιχείων", title: "Εξακριβώστε τον πραγματικό δικαιούχο", out: "Χαρτογράφηση δομής ιδιοκτησίας, αποδεικτικό ΚΜΠΔ του πελάτη", q: [7] },
  { n: 6, stage: "Συγκέντρωση στοιχείων", title: "Καταγράψτε σκοπό και φύση της σχέσης", out: "Σημείωμα ανά πελάτη, με αναφορά των πηγών", q: [23] },
  { n: 7, stage: "Αξιολόγηση κινδύνου", title: "Φτιάξτε τη μεθοδολογία βαθμολόγησης", out: "Γραπτή μέθοδος, στηριγμένη στα Παραρτήματα Ι και ΙΙ", q: [15, 17] },
  { n: 8, stage: "Αξιολόγηση κινδύνου", title: "Βαθμολογήστε κάθε πελάτη και κατατάξτε τον", out: "Φύλλο ανά πελάτη → είδος δέουσας επιμέλειας", q: [15] },
  { n: 9, stage: "Αξιολόγηση κινδύνου", title: "Τεκμηριώστε και συντάξτε την κατάσταση πελατών", out: "Μία γραμμή ανά πελάτη, με κατηγορία και ημερομηνία, έκδοση 1.0", q: [15, 16, 17] },
  { n: 10, stage: "Θεσμοθέτηση", title: "Εγχειρίδιο Πολιτικών και Διαδικασιών", out: "Εγκεκριμένο κείμενο, με ημερομηνία έγκρισης και επόμενης επανεξέτασης", q: [2] },
  { n: 11, stage: "Θεσμοθέτηση", title: "Δίαυλος καταγγελιών, καταλληλότητα, εσωτερικός έλεγχος", out: "Οι τρεις απαιτήσεις της παρ. 3 του άρθρου 35", q: [4] },
  { n: 12, stage: "Θεσμοθέτηση", title: "Εκπαίδευση και αρχείο εκπαίδευσης", out: "Πρόγραμμα με βάση τους κινδύνους, αποδεικτικά", q: [1] },
  { n: 13, stage: "Λειτουργία", title: "Εφαρμόστε τα μέτρα ανά κατηγορία κινδύνου", out: "Απλουστευμένη, συνήθης ή αυξημένη δέουσα επιμέλεια", q: [10] },
  { n: 14, stage: "Λειτουργία", title: "Παρακολούθηση και επικαιροποίηση", out: "Ημερομηνίες επανεξέτασης, έκτακτη επανεξέταση σε κάθε μεταβολή", q: [12, 13, 22] },
  { n: 15, stage: "Λειτουργία", title: "Διαδικασία ύποπτων συναλλαγών", out: "Γραπτή ακολουθία, αρχείο και για όσα εξετάστηκαν χωρίς αναφορά", q: [11, 18, 19] },
];
// ---- The 14 documents of the firm's compliance file.
const AML_FILE_DOCUMENTS = [
  { n: 1, title: "Πράξη ορισμού υπευθύνου συμμόρφωσης, με ημερομηνία και αρμοδιότητες", step: 1 },
  { n: 2, title: "Αποδεικτικό καταχώρισης του γραφείου στο ΚΜΠΔ, ειδικό μητρώο στην έδρα και Πίνακας Ι΄ του Ε3 κάθε έτους", step: 2 },
  { n: 3, title: "Φύλλο απογραφής πελατολογίου, με ημερομηνία επόμενης επανεξέτασης ανά πελάτη", step: 3 },
  { n: 4, title: "Έντυπο γνωριμίας (KYC) ανά πελάτη, με τα έγγραφα της ΠΟΛ.1200/2018 και την ενημέρωση GDPR", step: 4 },
  { n: 5, title: "Έντυπο πραγματικού δικαιούχου, χαρτογράφηση δομής και αποδεικτικό ΚΜΠΔ", step: 5 },
  { n: 6, title: "Σημείωμα σκοπού και φύσης της σχέσης, ή συμφωνητικό υπηρεσιών", step: 6 },
  { n: 7, title: "Γραπτή μεθοδολογία ανάλυσης κινδύνου (Κεφάλαιο 2 του εγχειριδίου) με το προφίλ του γραφείου", step: 7 },
  { n: 8, title: "Φύλλο ανάλυσης κινδύνου ανά πελάτη, χρονολογημένο και υπογεγραμμένο", step: 8 },
  { n: 9, title: "Συγκεντρωτική κατάσταση κατάταξης πελατών, με αριθμό έκδοσης και ημερομηνία", step: 9 },
  { n: 10, title: "Εγχειρίδιο Πολιτικών και Διαδικασιών, με ημερομηνία έγκρισης και επόμενης επανεξέτασης", step: 10 },
  { n: 11, title: "Δίαυλος καταγγελιών, διαδικασία καταλληλότητας προσωπικού και πόρισμα ετήσιου εσωτερικού ελέγχου", step: 11 },
  { n: 12, title: "Αρχείο εκπαίδευσης, με ημερομηνίες, θεματολογία και αποδεικτικά", step: 12 },
  { n: 13, title: "Γραπτές εγκρίσεις ανώτερου στελέχους για πελάτες υψηλού κινδύνου και ΠΕΠ", step: 13 },
  { n: 14, title: "Γραπτή διαδικασία ύποπτων συναλλαγών, αρχείο αναφορών, αρχείο περιστατικών χωρίς αναφορά, αρχείο απορρίψεων", step: 15 },
];
const AML_CASE_DECISIONS = { pending: "Υπό εξέταση", reported: "Αναφέρθηκε στην Αρχή", not_reported: "Εξετάστηκε — δεν αναφέρθηκε", abstained: "Αποχή από συναλλαγή (άρ. 23)" };
const AML_CASE_TONE = { pending: "danger", reported: "accent", not_reported: "neutral", abstained: "warn" };
const AML_AUTHORITIES = ["ΑΑΔΕ (εποπτική αρχή)", "Αρχή Καταπολέμησης Νομιμοποίησης Εσόδων", "Υπουργείο Οικονομικών", "Εισαγγελία / Δικαστήριο", "Άλλη δημόσια αρχή"];
const mapAmlCase = (r) => ({ id: r.id, clientId: r.client_id || null, detectedOn: r.detected_on, detectedBy: r.detected_by || null, description: r.description, decision: r.decision, decisionOn: r.decision_on || "", decidedBy: r.decided_by || null, reportRef: r.report_ref || "", rationale: r.rationale || "", createdAt: r.created_at });
const mapAuthorityRequest = (r) => ({ id: r.id, authority: r.authority, receivedOn: r.received_on, dueOn: r.due_on || "", subject: r.subject, clientId: r.client_id || null, status: r.status, answeredOn: r.answered_on || "", note: r.note || "" });
const DEFAULT_AML_FIRM = {
  complianceOfficerId: "", appointmentDate: "", kmpdRegisteredOn: "", kmpdLastChangeOn: "", kmpdLastFiledOn: "",
  e3PanelIYears: [], manualVersion: "", manualApprovedOn: "", manualNextReviewOn: "", lastTrainingOn: "", lastInternalAuditOn: "",
  registerVersion: "1.0", registerDate: "", registerApprovedBy: "", methodologyModel: "A",
  selfAssessment: {}, steps: {}, fileDocs: {},
};
const addMonths = (d, m) => { const r = new Date(d); r.setMonth(r.getMonth() + m); return r; };
const daysUntil = (key) => key ? Math.round((fromKey(key) - TODAY) / 86400000) : null;

// ---- Alerts engine: every reporting / review / filing obligation that is
// due, with a due date and severity. Pure function, used by the Compliance
// page and the dashboard.
function amlAlerts({ clients, amlRecords, cases, requests, amlFirm, employees }) {
  const out = [];
  const push = (severity, kind, title, detail, dueKey, clientId) => out.push({ severity, kind, title, detail, due: dueKey || null, clientId: clientId || null });
  const year = TODAY.getFullYear();
  const active = clients.filter((c) => c.lifecycle === "live" || c.lifecycle === "onboarding");
  // Client files
  active.forEach((c) => {
    const a = amlRecords.find((x) => x.clientId === c.id);
    if (!a) { push(c.lifecycle === "live" ? "high" : "medium", "file", `Χωρίς φάκελο δέουσας επιμέλειας: ${c.name}`, c.lifecycle === "live" ? "Υφιστάμενος πελάτης χωρίς KYC και αξιολόγηση κινδύνου (άρ. 13 παρ. 7)" : "Ταυτοποίηση πριν την έναρξη της σχέσης (άρ. 14 παρ. 1)", null, c.id); return; }
    if (a.status === "declined") return;
    if (c.lifecycle === "live" && a.status !== "completed") push("high", "file", `Ανολοκλήρωτος φάκελος: ${c.name}`, "Ο πελάτης είναι live ενώ η δέουσα επιμέλεια δεν έχει ολοκληρωθεί (άρ. 14 παρ. 1-2)", null, c.id);
    if (a.riskLevel === "high" && !a.approvedBy) push("high", "approval", `Εκκρεμεί έγκριση ανώτερου στελέχους: ${c.name}`, "Υψηλός κίνδυνος / ΠΕΠ χωρίς γραπτή έγκριση (άρ. 16Α, 18)", null, c.id);
    if (a.nextReviewOn) {
      const d = daysUntil(a.nextReviewOn);
      if (d < 0) push("high", "review", `Επανεξέταση σε καθυστέρηση: ${c.name}`, `Έπρεπε ${a.nextReviewOn} (${AML_LEVEL_LABELS[a.riskLevel]} κίνδυνος)`, a.nextReviewOn, c.id);
      else if (d <= 30) push("medium", "review", `Επανεξέταση εντός 30 ημερών: ${c.name}`, `Προγραμματισμένη ${a.nextReviewOn}`, a.nextReviewOn, c.id);
    } else if (a.status === "completed") push("medium", "review", `Χωρίς ημερομηνία επανεξέτασης: ${c.name}`, "Ορίστε επόμενη επικαιροποίηση (Βήμα 14)", null, c.id);
    if (a.uboChangeDate && !a.uboFiledDate) {
      const deadline = toKey(addDays(fromKey(a.uboChangeDate), 60));
      const d = daysUntil(deadline);
      push(d < 0 ? "high" : "medium", "kmpd", `Μεταβολή πραγματικού δικαιούχου — ΚΜΠΔ: ${c.name}`, `Καταχώριση εντός 60 ημερών από ${a.uboChangeDate} (άρ. 20 παρ. 1) — προθεσμία ${deadline}${d < 0 ? ", εκπρόθεσμη" : ""}`, deadline, c.id);
    }
    if (a.pepStatus !== "none" && a.riskLevel !== "high") push("high", "file", `ΠΕΠ χωρίς κατάταξη σε υψηλό κίνδυνο: ${c.name}`, "Το άρθρο 18 υπερισχύει — ελέγξτε την αξιολόγηση", null, c.id);
  });
  // Cases
  (cases || []).filter((k) => k.decision === "pending").forEach((k) => {
    const d = -daysUntil(k.detectedOn);
    push(d > 7 ? "high" : "medium", "case", `Περιστατικό υπό εξέταση${k.clientId ? `: ${clients.find((c) => c.id === k.clientId)?.name || ""}` : ""}`, `Εντοπίστηκε ${k.detectedOn} — απόφαση αναφοράς «αμελλητί» (άρ. 22). Μην ενημερώσετε τον πελάτη (άρ. 27).`, k.detectedOn, k.clientId);
  });
  // Authority requests
  (requests || []).filter((r) => r.status === "open").forEach((r) => {
    const d = daysUntil(r.dueOn);
    push(d !== null && d < 0 ? "high" : d !== null && d <= 7 ? "high" : "medium", "request", `Αίτημα αρχής: ${r.authority}`, `${r.subject}${r.dueOn ? ` — προθεσμία ${r.dueOn}` : " — χωρίς προθεσμία, απάντηση «αμελλητί» (άρ. 22 παρ. 1β)"}`, r.dueOn || null, r.clientId);
  });
  // Firm-level
  const f = { ...DEFAULT_AML_FIRM, ...(amlFirm || {}) };
  if (!f.complianceOfficerId) push("high", "firm", "Δεν έχει οριστεί υπεύθυνος συμμόρφωσης", "Γραπτή πράξη ορισμού (άρ. 35 παρ. 3α, ερώτημα 3)");
  if (!f.kmpdRegisteredOn) push("high", "firm", "Το γραφείο δεν έχει καταχωριστεί στο ΚΜΠΔ", "Άρ. 20 — ερώτημα 20 του ελέγχου");
  if (f.kmpdLastChangeOn && !f.kmpdLastFiledOn) { const dl = toKey(addDays(fromKey(f.kmpdLastChangeOn), 60)); push(daysUntil(dl) < 0 ? "high" : "medium", "firm", "Μεταβολή πραγματικών δικαιούχων του γραφείου — ΚΜΠΔ", `Καταχώριση έως ${dl} (60 ημέρες, άρ. 20 παρ. 1)`, dl); }
  if (!(f.e3PanelIYears || []).includes(year - 1)) push(TODAY.getMonth() >= 3 ? "high" : "medium", "firm", `Πίνακας Ι΄ του Ε3 για το φορολογικό έτος ${year - 1}`, "Συμπληρώνεται με τη δήλωση εισοδήματος του γραφείου (ερώτημα 21)", `${year}-07-31`);
  if (!f.manualApprovedOn) push("high", "firm", "Χωρίς εγκεκριμένο Εγχειρίδιο Πολιτικών και Διαδικασιών", "Βήμα 10 — ερώτημα 2");
  else if (f.manualNextReviewOn && daysUntil(f.manualNextReviewOn) < 30) push(daysUntil(f.manualNextReviewOn) < 0 ? "high" : "medium", "firm", "Επανεξέταση του Εγχειριδίου", `Προγραμματισμένη ${f.manualNextReviewOn}`, f.manualNextReviewOn);
  if (!f.lastTrainingOn || daysUntil(toKey(addMonths(fromKey(f.lastTrainingOn), 12))) < 0) push("medium", "firm", "Ετήσια εκπαίδευση προσωπικού", f.lastTrainingOn ? `Τελευταία ${f.lastTrainingOn} — πάνω από 12 μήνες (άρ. 37)` : "Δεν έχει καταγραφεί εκπαίδευση (ερώτημα 1)", f.lastTrainingOn ? toKey(addMonths(fromKey(f.lastTrainingOn), 12)) : null);
  if (!f.lastInternalAuditOn || daysUntil(toKey(addMonths(fromKey(f.lastInternalAuditOn), 12))) < 0) push("medium", "firm", "Ετήσιος εσωτερικός έλεγχος / αυτοαξιολόγηση", f.lastInternalAuditOn ? `Τελευταίο πόρισμα ${f.lastInternalAuditOn}` : "Δεν έχει καταγραφεί πόρισμα (άρ. 35 παρ. 3γ, ερώτημα 4)", f.lastInternalAuditOn ? toKey(addMonths(fromKey(f.lastInternalAuditOn), 12)) : null);
  const newerThanRegister = f.registerDate ? amlRecords.filter((a) => a.assessmentDate && a.assessmentDate > f.registerDate).length : amlRecords.length;
  if (newerThanRegister > 0) push("medium", "register", "Νέα έκδοση συγκεντρωτικής κατάστασης πελατών", `${newerThanRegister} αξιολόγηση(-εις) μετά την έκδοση ${f.registerVersion || "—"}${f.registerDate ? ` της ${f.registerDate}` : ""} (άρ. 35 παρ. 2, Βήμα 9)`);
  const sevRank = { high: 0, medium: 1, low: 2 };
  return out.sort((x, y) => sevRank[x.severity] - sevRank[y.severity] || (x.due || "9999").localeCompare(y.due || "9999"));
}
const mapContact = (r) => ({ id: r.id, clientId: r.client_id, name: r.name, role: r.role || "", email: r.email || "", phone: r.phone || "", note: r.note || "", isPrimary: !!r.is_primary });
const mapNote = (r) => ({ id: r.id, clientId: r.client_id, authorId: r.author_id || null, kind: r.kind, noteDate: r.note_date, content: r.content, createdAt: r.created_at });
const mapTask = (r) => ({ id: r.id, clientId: r.client_id || null, kind: r.kind || "general", stage: r.stage || null, title: r.title, description: r.description || "", assignedTo: r.assigned_to || null, createdBy: r.created_by || null, dueDate: r.due_date || null, priority: r.priority, status: r.status, doneAt: r.done_at || null, createdAt: r.created_at });

// CRM vocab
const LEGAL_FORMS = ["Ατομική", "Ο.Ε.", "Ε.Ε.", "Ι.Κ.Ε.", "Ε.Π.Ε.", "Α.Ε.", "Υποκατάστημα αλλοδαπής", "Ιδιώτης", "Άλλο"];
const NOTE_KINDS = { call: "Call", email: "Email", meeting: "Meeting", note: "Note" };
const TASK_PRIORITIES = { high: "High", normal: "Normal", low: "Low" };
const TASK_PRIORITY_TONE = { high: "danger", normal: "neutral", low: "accent" };
const TASK_STATUSES = { open: "Open", in_progress: "In progress", done: "Done" };
const TASK_STATUS_TONE = { open: "neutral", in_progress: "warn", done: "accent" };
// Open tasks: overdue if the due date has passed.
const taskIsOverdue = (t) => t.status !== "done" && t.dueDate && fromKey(t.dueDate) < TODAY;
const LIFECYCLE_LABELS = { lead: "Lead", onboarding: "Onboarding", live: "Live", ended: "Ended" };
const LIFECYCLE_TONE = { lead: "neutral", onboarding: "warn", live: "accent", ended: "danger" };
// Onboarding progress for a client: its checklist tasks and how many are done.
function onboardingProgress(client, tasks) {
  const list = tasks.filter((t) => t.clientId === client.id && t.kind === "onboarding");
  const done = list.filter((t) => t.status === "done").length;
  return { list, done, total: list.length, complete: list.length > 0 && done === list.length };
}
// Flowchart of the client lifecycle, shown from the ⓘ button on the client card.
function OnboardingFlowchart() {
  const box = (x, y, w, h, title, sub, fill, stroke) => (
    <g>
      <rect x={x} y={y} width={w} height={h} rx={9} fill={fill || C.surface} stroke={stroke || C.border} strokeWidth="1.5" />
      <text x={x + w / 2} y={y + (sub ? h / 2 - 5 : h / 2 + 4)} textAnchor="middle" fontFamily={sans} fontSize="12.5" fontWeight="700" fill={C.ink}>{title}</text>
      {sub && <text x={x + w / 2} y={y + h / 2 + 12} textAnchor="middle" fontFamily={sans} fontSize="10.5" fill={C.inkMuted}>{sub}</text>}
    </g>
  );
  const arrow = (x1, y1, x2, y2) => <line x1={x1} y1={y1} x2={x2} y2={y2} stroke={C.inkFaint} strokeWidth="1.6" markerEnd="url(#fc-arrow)" />;
  return (
    <div>
      <svg viewBox="0 0 640 660" style={{ width: "100%", height: "auto" }} role="img" aria-label="Client lifecycle flowchart">
        <defs><marker id="fc-arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0 0 L10 5 L0 10 z" fill={C.inkFaint} /></marker></defs>
        {box(180, 10, 280, 50, "Lead (prospect)", "Stage: Lead — no tasks yet", C.surface)}
        {arrow(320, 60, 320, 88)}
        {box(160, 90, 320, 56, "Start onboarding", "Checklist created · set the engagement date", C.surface, C.warn)}
        {arrow(320, 146, 320, 174)}
        {box(120, 176, 400, 56, "1 · Due diligence (ν. 4557/2018)", "KYC, UBOs, risk, PEP — due on the engagement date", C.surface)}
        {arrow(320, 232, 320, 260)}
        {box(120, 262, 400, 56, "2 · Contract & authorisations", "Agreement signed · ΑΑΔΕ/ΕΦΚΑ — due on the engagement date", C.surface)}
        {arrow(320, 318, 320, 346)}
        {box(120, 348, 400, 56, "3 · Company formation (new entity only)", "ΥΜΣ, ΓΕΜΗ, bank, ΕΦΚΑ — 3 working days after", C.surface)}
        {arrow(320, 404, 320, 432)}
        {box(120, 434, 400, 70, "4 · Client setup", "Hyper · tax system · e-invoicing · providers · POS", C.surface)}
        <text x="320" y="494" textAnchor="middle" fontFamily={sans} fontSize="10.5" fill={C.inkMuted}>due 5 working days after the engagement date</text>
        {arrow(320, 504, 320, 532)}
        <path d="M320 534 L470 586 L320 638 L170 586 Z" fill={C.surface} stroke={C.accent} strokeWidth="1.8" />
        <text x="320" y="578" textAnchor="middle" fontFamily={sans} fontSize="12" fontWeight="700" fill={C.ink}>Ready to go live?</text>
        <text x="320" y="595" textAnchor="middle" fontFamily={sans} fontSize="10" fill={C.inkMuted}>Fees set · Agreement signed</text>
        <text x="320" y="608" textAnchor="middle" fontFamily={sans} fontSize="10" fill={C.inkMuted}>DD completed · All tasks done</text>
        {arrow(470, 586, 540, 586)}
        <text x="505" y="578" textAnchor="middle" fontFamily={sans} fontSize="10.5" fontWeight="700" fill={C.accent}>Yes</text>
        {box(542, 560, 92, 52, "LIVE", "go-live date set", C.surface, C.accent)}
        <path d="M170 586 L40 586 L40 205 L118 205" fill="none" stroke={C.warn} strokeWidth="1.6" strokeDasharray="5 4" markerEnd="url(#fc-arrow)" />
        <text x="26" y="400" textAnchor="middle" fontFamily={sans} fontSize="10.5" fontWeight="700" fill={C.warn} transform="rotate(-90 26 400)">No — finish what is missing</text>
      </svg>
      <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginTop: 6 }}>
        The engagement date is the signing date of the agreement (or the date set in Details until it is signed). Deadlines count working days, excluding weekends and Greek public holidays.
      </div>
    </div>
  );
}
// Go-live readiness: a client goes live only when all four are in place —
// fees set, agreement signed, due diligence completed, every onboarding task done.
function goLiveReadiness(client, tasks, aml, documents = []) {
  const prog = onboardingProgress(client, tasks);
  const checks = [
    { key: "fees", label: "Fees & hours — fixed fee set", ok: !!currentFeeOf(client).fixedFee },
    { key: "agreement", label: "Agreement — signed", ok: client.contractStatus === "signed" || documents.some((d) => d.clientId === client.id && d.category === "contract") },
    { key: "dd", label: "Due diligence — completed", ok: aml?.status === "completed" },
    { key: "tasks", label: `Tasks — all done (${prog.done}/${prog.total})`, ok: prog.complete },
  ];
  return { checks, ready: checks.every((c) => c.ok) };
}
// Starts onboarding: creates one task per active checklist item (assigned
// to the account manager, or whoever starts it) and moves the client to
// 'onboarding'. Items already present for this client are not duplicated.
// includeFormation: also create the "Company formation" tasks (for a newly
// formed entity). Items marked for capital companies are skipped unless the
// client's legal form is Α.Ε./Ε.Π.Ε./Ι.Κ.Ε.; "if applicable" items are
// created with that note so they can be closed or deleted quickly.
// Onboarding deadlines are counted from the engagement (contract) date:
// due diligence and contract & authorisations are due on that date, company
// formation 3 working days later, client setup 5 working days later.
const ONBOARDING_OFFSET_WD = { aml: 0, contract: 0, formation: 3, setup: 5 };
function addWorkingDays(d, n) {
  const hol = holidayMapForYears([d.getFullYear(), d.getFullYear() + 1]);
  let cur = new Date(d), left = n;
  while (left > 0) { cur = addDays(cur, 1); if (!isWeekend(cur) && !hol.has(toKey(cur))) left--; }
  return cur;
}
const onboardingBaseDate = (c) => c?.contractSignedOn || c?.engagementDate || null;
function onboardingDueFor(stage, baseKey) {
  return toKey(addWorkingDays(fromKey(baseKey), ONBOARDING_OFFSET_WD[stage] ?? 5));
}
// Re-dates every open onboarding task of the client from the base date.
async function applyOnboardingDeadlines(clientId, baseKey) {
  if (!baseKey) return;
  for (const stage of Object.keys(ONBOARDING_OFFSET_WD)) {
    await supabase.from("tasks").update({ due_date: onboardingDueFor(stage, baseKey) }).eq("client_id", clientId).eq("kind", "onboarding").eq("stage", stage).neq("status", "done");
  }
}
async function startOnboarding(client, templates, existingTasks, user, { includeFormation = false } = {}) {
  const have = new Set(existingTasks.filter((t) => t.clientId === client.id && t.kind === "onboarding").map((t) => t.title));
  const stageRank = (k) => ONBOARDING_STAGES.findIndex((s) => s.key === k);
  const applicable = (tpl) => {
    if (tpl.stage === "formation" && !includeFormation) return false;
    if (tpl.appliesTo === "capital" && !isCapitalCompany(client.legalForm)) return false;
    return true;
  };
  const rows = templates.filter((tpl) => tpl.active && !have.has(tpl.title) && applicable(tpl)).sort((a, b) => stageRank(a.stage) - stageRank(b.stage) || a.sortOrder - b.sortOrder).map((tpl) => ({
    client_id: client.id, kind: "onboarding", stage: tpl.stage, title: tpl.title, description: [tpl.appliesTo === "if_applicable" ? "If applicable — close or delete if not relevant." : null, tpl.description || null].filter(Boolean).join(" ") || null,
    assigned_to: client.accountManagerId || user.id, created_by: user.id,
    due_date: onboardingBaseDate(client) ? onboardingDueFor(tpl.stage, onboardingBaseDate(client)) : toKey(addDays(TODAY, tpl.dueDays)), priority: "normal", status: "open",
  }));
  if (rows.length) {
    const { error } = await supabase.from("tasks").insert(rows);
    if (error) { window.alert(`Couldn't create onboarding tasks: ${error.message}`); return false; }
  }
  const { error } = await supabase.from("clients").update({ lifecycle: "onboarding", go_live_date: null }).eq("id", client.id);
  if (error) { window.alert(`Couldn't update client: ${error.message}`); return false; }
  return true;
}
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
const mapBillingStatus = (r) => ({
  id: r.id, clientId: r.client_id, kind: r.kind, month: r.month, extraFeeId: r.extra_fee_id || null,
  invoicedDate: r.invoiced_date || null, invoiceRef: r.invoice_ref || "",
  paidDate: r.paid_date || null, paidAmount: r.paid_amount === null || r.paid_amount === undefined ? null : Number(r.paid_amount),
  note: r.note || "",
});
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

function EmployeeTimesheet({ user, clients, entries, refetchEntries, lockedWeeks, refetchLockedWeeks, leaveRequests = [] }) {
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
  // Leave from the Time Off module is shown on the day (no hours are created). Approved leave on a working day lowers the week's target.
  const myLeave = leaveRequests.filter((r) => r.employeeId === user.id && r.status !== "rejected");
  const weekHolidays = holidayMapForYears([weekStart.getFullYear(), addDays(weekStart, 6).getFullYear()]);
  const leaveOnDay = (key) => {
    const hits = myLeave.filter((r) => r.startDate <= key && key <= r.endDate);
    return hits.find((r) => r.status === "approved") || hits[0] || null;
  };
  const isWorkingDay = (d) => !isWeekend(d) && !weekHolidays.has(toKey(d));
  const approvedLeaveDays = days.filter((d) => isWorkingDay(d) && leaveOnDay(toKey(d))?.status === "approved").length;
  const baseWeeklyTarget = user.weeklyHours || 40;
  const weeklyTarget = Math.max(0, Math.round(baseWeeklyTarget * (1 - approvedLeaveDays / 5) * 100) / 100);
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
        <StatCard label={`vs. ${weeklyTarget}h/week target${approvedLeaveDays ? ` (${approvedLeaveDays} leave day${approvedLeaveDays > 1 ? "s" : ""})` : ""}`} value={weeklyDiffLabel} accent={weeklyDiffColor} />
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
          const dayLeave = !isWeekend(day) ? leaveOnDay(toKey(day)) : null;
          const dayHoliday = weekHolidays.get(toKey(day));
          return (
            <div key={toKey(day)} style={{ display: "flex", gap: 16, padding: "10px 0", borderBottom: `1px solid ${C.border}`, background: dayLeave ? C.accentSoft : "transparent" }}>
              <div style={{ width: 90, flexShrink: 0 }}>
                <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 600, color: isToday(day) ? C.accent : C.ink }}>
                  {fmtDow(day)} {fmtShort(day)}
                </div>
                {isWeekend(day) && <div style={{ fontFamily: sans, fontSize: 10.5, color: C.inkFaint }}>Weekend</div>}
                {dayHoliday && !isWeekend(day) && <div style={{ fontFamily: sans, fontSize: 10.5, color: C.inkFaint }}>{dayHoliday}</div>}
                {!isLocked && (
                  <button onClick={() => setForm((f) => ({ ...f, date: toKey(day) }))} title={`Add hours for ${fmtShort(day)}`}
                    style={{ ...iconBtnStyle, padding: 0, marginTop: 4, fontFamily: sans, fontSize: 11, color: form.date === toKey(day) ? C.accent : C.inkFaint, fontWeight: form.date === toKey(day) ? 700 : 500, alignItems: "center", gap: 3 }}>
                    <Plus size={11} /> {form.date === toKey(day) ? "Selected" : "Add hours"}
                  </button>
                )}
              </div>
              <div style={{ flex: 1 }}>
                {dayLeave && (
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: dayEntries.length ? 8 : 0 }}>
                    <Badge tone={dayLeave.status === "approved" ? "accent" : "warn"}>{leaveTypeLabel(dayLeave.type)} · {dayLeave.status === "approved" ? "Approved" : "Pending approval"}</Badge>
                    {dayEntries.length === 0 && <span style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted }}>On leave</span>}
                  </div>
                )}
                {dayEntries.length === 0 ? (
                  dayLeave ? null : <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkFaint }}>No hours logged</div>
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

function AdminDashboard({ employees, clients, entries, rentalProperties, rentalExpenses, operatingExpenses = [], tasks = [], onOpenCard, templates = [], complianceAlerts = [], onOpenCompliance, billingStatuses = [], onOpenBilling, amlRecords = [] }) {
  const [period, setPeriod] = useState("month"); // "week" | "month" | "quarter" | "year"
  const [anchor, setAnchor] = useState(TODAY);
  const [lossModal, setLossModal] = useState(false);
  const DEFAULT_DASHBOARD_CONFIG = {
    statCards: true, billingPanel: true, ddPanel: true, compliancePanel: true, onboardingPanel: true, tasksPanel: true, growthPanel: true, allocationChart: true, revenueCostChart: true,
    byClientTable: true, byEmployeeTable: true, nonChargeableTable: true, profitabilityTable: true, departmentTable: true, costCenterTable: true,
    monthlyEmployeeTable: true, weeklyEmployeeTable: true, monthlyClientTable: true, companyMonthlyTable: true, effectiveRateTable: true, clientProfitTable: true,
  };
  const [configRaw, setConfig] = useAppSetting("dashboard_config", DEFAULT_DASHBOARD_CONFIG);
  // Merge with the defaults so panels added later (billing, compliance, tasks…) appear even when a layout was saved earlier.
  const config = { ...DEFAULT_DASHBOARD_CONFIG, ...(configRaw || {}) };
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
  // Rental expenses are classified as operating expenses (below gross profit), so
  // cost / gross profit = labor only.
  const totalCost = totalLaborCost;
  const totalProfit = totalRevenue - totalCost;
  const overallMargin = totalRevenue > 0 ? (totalProfit / totalRevenue) * 100 : null;
  // Other operating expenses (insurance, software, bank charges, depreciation, AI, business development…):
  // gross profit (above) minus these = operating profit.
  const opex = opexForRange(operatingExpenses, rangeStart, rangeEnd, rentalExpenses);
  const operatingProfit = totalProfit - opex.total;
  const operatingMargin = totalRevenue > 0 ? (operatingProfit / totalRevenue) * 100 : null;

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
  const unallocatedCostByDepartment = Object.fromEntries(DEPARTMENTS.map((d) => [d.key, 0]));
  employeeCosts.forEach((ec) => {
    costByDepartment[ec.department] = (costByDepartment[ec.department] || 0) + ec.cost;
    unallocatedCostByDepartment[ec.department] = (unallocatedCostByDepartment[ec.department] || 0) + (ec.byCostCenter[UNALLOCATED] || 0);
  });

  // Profitability by client — gross margin and operating margin.
  //  • Gross margin = (revenue − labor cost of the hours spent on the client) ÷ revenue.
  //  • Operating margin also charges each client its share of the OVERHEAD, allocated by cost
  //    center with the client's HOURS as the driver:
  //      pool of a cost center = labor cost not charged to clients (internal activities, hours not
  //                              logged on a client) + operating expenses of that cost center
  //      client's share        = pool × client hours in that cost center ÷ all client hours in it
  //    Accounting / Tax / Payroll / Other use the hours logged with that category; Director,
  //    Finance, Rental and General (company-wide) use the client's total hours. The rows add up
  //    to the company's operating profit.
  const OVERHEAD_CCS = [...COST_CENTERS, "general"];
  const clientIdsInRange = new Set(byClient.map((c) => c.id));
  const billed = byClient.filter((c) => c.revenue > 0.005);
  const ovhDriver = Object.fromEntries(OVERHEAD_CCS.map((cc) => [cc, {}]));
  billed.forEach((c) => {
    const ents = rangeEntries.filter((e) => e.clientId === c.id);
    const total = ents.reduce((s, e) => s + e.hours, 0);
    OVERHEAD_CCS.forEach((cc) => { ovhDriver[cc][c.id] = CATEGORIES.includes(cc) ? ents.filter((e) => (CATEGORIES.includes(e.category) ? e.category : "other") === cc).reduce((s, e) => s + e.hours, 0) : total; });
  });
  const ovhDriverTotal = (cc) => Object.values(ovhDriver[cc]).reduce((s, v) => s + v, 0);
  const ovhPool = Object.fromEntries(OVERHEAD_CCS.map((cc) => [cc, 0]));
  employeeCosts.forEach((ec) => {
    if (!(ec.cost > 0)) return;
    const attributed = ec.hours > 0 ? ec.cost * (ec.entries.filter((e) => e.clientId && clientIdsInRange.has(e.clientId)).reduce((s, e) => s + e.hours, 0) / ec.hours) : 0;
    const unattributed = ec.cost - attributed;
    Object.entries(ec.byCostCenter).forEach(([cc, v]) => { ovhPool[cc === UNALLOCATED ? "general" : cc] += unattributed * (v / ec.cost); });
  });
  const spreadPool = (amount, keys, weightOf) => {
    const w = keys.map(weightOf), t = w.reduce((s, x) => s + x, 0);
    keys.forEach((k, i) => { ovhPool[k] += t > 0 ? amount * (w[i] / t) : amount / keys.length; });
  };
  spreadPool(opex.byDepartment.accounting_team || 0, CATEGORIES, (cc) => ovhDriverTotal(cc));
  spreadPool(opex.byDepartment.management_team || 0, ["director", "finance"], (cc) => revenueByCC[cc] || 0);
  ovhPool.rental += opex.byDepartment.rental || 0;
  ovhPool.general += opex.byDepartment[OPEX_GENERAL] || 0;
  const clientOverhead = Object.fromEntries(billed.map((c) => [c.id, 0]));
  const overheadByCC = OVERHEAD_CCS.map((cc) => {
    const pool = ovhPool[cc], hours = ovhDriverTotal(cc);
    billed.forEach((c) => { clientOverhead[c.id] += hours > 0 ? pool * (ovhDriver[cc][c.id] / hours) : pool / billed.length; });
    return { key: cc, label: cc === "general" ? "General (company-wide)" : costCenterLabel(cc), pool, hours, perHour: hours > 0 ? pool / hours : null };
  }).filter((r) => Math.abs(r.pool) > 0.005);
  const clientProfitRows = billed.map((c) => {
    const grossProfit = c.revenue - c.cost, overhead = clientOverhead[c.id], operatingProfitC = grossProfit - overhead;
    return { id: c.id, name: c.name, hours: c.hours, revenue: c.revenue, cost: c.cost, grossProfit, grossMargin: (grossProfit / c.revenue) * 100, overhead, operatingProfit: operatingProfitC, operatingMargin: (operatingProfitC / c.revenue) * 100 };
  }).sort((a, b) => a.operatingMargin - b.operatingMargin);
  // Clients with negative operating profit in the selected period — shown in a pop-up.
  const lossClients = clientProfitRows.filter((r) => r.operatingProfit < -0.005).sort((a, b) => a.operatingProfit - b.operatingProfit);
  const lossTotal = lossClients.reduce((s, r) => s + r.operatingProfit, 0);
  useEffect(() => {
    // Opens by itself once per browser session when the dashboard loads.
    if (lossClients.length === 0) return;
    let seen = false;
    try { seen = sessionStorage.getItem("gloria_loss_popup") === "1"; } catch { /* ignore */ }
    if (!seen) { setLossModal(true); try { sessionStorage.setItem("gloria_loss_popup", "1"); } catch { /* ignore */ } }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lossClients.length > 0]);

  // Revenue, labor cost and profit per cost center for the period. The
  // rows add up to the same totals as the cards above; "Unallocated" only
  // appears when some fixed fee has no split yet or a department's cost had
  // no hours or revenue to spread it by.
  const byCostCenter = [...COST_CENTERS, UNALLOCATED].map((cc) => {
    const revenue = revenueByCC[cc] || 0;
    const cost = employeeCosts.reduce((s, ec) => s + (ec.byCostCenter?.[cc] || 0), 0);
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

  // Effective rate by client and month (revenue ÷ hours); cost/h allocates each
  // employee's monthly cost by their share of hours on the client that month.
  const effRateRows = clientMonthly.map((c) => {
    let hours = 0, revenue = 0, cost = 0;
    const cells = trendMonths.map((m, i) => {
      const monthAll = entries.filter((e) => fromKey(e.date) >= m.start && fromKey(e.date) <= m.end);
      const cE = monthAll.filter((e) => e.clientId === c.id);
      let mc = 0;
      employees.forEach((emp) => {
        const mine = cE.filter((e) => e.employeeId === emp.id).reduce((s, e) => s + e.hours, 0);
        if (!mine) return;
        const all = monthAll.filter((e) => e.employeeId === emp.id).reduce((s, e) => s + e.hours, 0);
        if (all > 0) mc += periodEmployeeCost(emp, m.start, m.end) * (mine / all);
      });
      const h = c.months[i].hours, r = c.months[i].revenue;
      hours += h; revenue += r; cost += mc;
      return { hours: h, rate: h > 0 ? r / h : null, costH: h > 0 ? mc / h : null };
    });
    return { id: c.id, name: c.name, cells, hours, rate: hours > 0 ? revenue / hours : null, costH: hours > 0 ? cost / hours : null };
  }).filter((r) => r.hours > 0).sort((a, b) => (a.rate ?? 1e9) - (b.rate ?? 1e9));

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
    const cost = employees.filter((e) => employeeBilledInRange(e, m.start, m.end)).reduce((sum, e) => sum + periodEmployeeCost(e, m.start, m.end), 0);
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
    statCards: "Summary cards", billingPanel: "Billing — outstanding alert", ddPanel: "Due diligence — client files alert", compliancePanel: "AML compliance alerts (firm, reports, authorities)", onboardingPanel: "Pipeline — leads & onboarding", tasksPanel: "Tasks — overdue & due this week", growthPanel: "Growth — vs. previous period and YoY", allocationChart: "Actual vs. allocated chart", revenueCostChart: "Revenue vs. cost chart",
    byClientTable: "Hours by client table", byEmployeeTable: "Hours by employee table", nonChargeableTable: "Non-chargeable time", profitabilityTable: "Profitability by department", departmentTable: "P&L by department (detail)", costCenterTable: "Revenue & cost by cost center",
    monthlyEmployeeTable: "Hours per employee, by month", weeklyEmployeeTable: "Hours per employee, by week", monthlyClientTable: "Client analysis, by month", effectiveRateTable: "Effective rate by client, by month", clientProfitTable: "Profitability by client (gross / operating margin)",
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

      {lossClients.length > 0 && (
        <div onClick={() => setLossModal(true)} style={{ cursor: "pointer", display: "flex", alignItems: "center", gap: 10, background: C.dangerSoft || "#fdecec", border: `1px solid ${C.danger}`, borderRadius: 8, padding: "9px 14px", marginBottom: 18, fontFamily: sans, fontSize: 13, color: C.danger, fontWeight: 600 }}>
          <AlertTriangle size={16} /> {lossClients.length} {lossClients.length === 1 ? "client has" : "clients have"} negative operating profit in this period ({fmtEur(lossTotal)}) — click to see which
        </div>
      )}
      {lossModal && (
        <Modal title={`Clients with negative operating profit — ${periodLabel(period, rangeStart, rangeEnd)}`} onClose={() => setLossModal(false)} width={760}>
          {lossClients.length === 0 ? <div style={{ fontFamily: sans, fontSize: 13, color: C.inkMuted }}>No client has negative operating profit in this period.</div> : (
            <>
              <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 10 }}>Operating profit = revenue − labor cost of the client's hours − allocated overhead. Total loss: <strong style={{ color: C.danger }}>{fmtEur(lossTotal)}</strong>. Change the period on the dashboard to see another month.</div>
              <TableShell headers={["Client", "Hours", "Revenue", "Gross profit", "Overhead", "Operating profit", "Margin"]}>
                {lossClients.map((r) => (
                  <tr key={r.id} style={{ cursor: "pointer" }} onClick={() => { setLossModal(false); onOpenCard && onOpenCard(r.id); }}>
                    <Td>{r.name}</Td>
                    <Td mono>{r.hours.toFixed(1)}</Td>
                    <Td mono>{fmtEur(r.revenue)}</Td>
                    <Td mono style={{ color: r.grossProfit < 0 ? C.danger : undefined }}>{fmtEur(r.grossProfit)}</Td>
                    <Td mono>{fmtEur(r.overhead)}</Td>
                    <Td mono style={{ fontWeight: 700, color: C.danger }}>{fmtEur(r.operatingProfit)}</Td>
                    <Td mono style={{ color: C.danger }}>{r.operatingMargin.toFixed(0)}%</Td>
                  </tr>
                ))}
              </TableShell>
              <div style={{ fontFamily: sans, fontSize: 11.5, color: C.inkFaint, marginTop: 8 }}>Click a client to open its card.</div>
            </>
          )}
        </Modal>
      )}

      {config.statCards && (
        <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
          <StatCard label="Active employees" value={activeEmployees} sub={`of ${employees.filter(e=>e.role==='employee' && (!earliestHistoryDate(e.costHistory) || earliestHistoryDate(e.costHistory) <= toKey(rangeEnd))).length} employees`} />
          <StatCard label="Clients with hours" value={byClient.filter((c) => c.hours > 0).length} sub={`of ${byClient.length} active clients`} />
          <StatCard label="New clients this period" value={newClientsThisPeriod.length} sub={newClientsThisPeriod.length ? newClientsThisPeriod.map(c=>c.name).join(", ") : "None started this period"} />
          <StatCard label="New employees this period" value={newEmployeesThisPeriod.length} sub={newEmployeesThisPeriod.length ? fmtEur(newHiresCost) + " added to labor cost" : "None started this period"} />
          <StatCard label="Revenue this period" value={fmtEur(totalRevenue)} accent={C.accent} sub="Fixed fees + extra fees" />
          <StatCard label="Labor cost this period" value={fmtEur(totalLaborCost)} sub="Gross + benefits, prorated" />
          <StatCard label="Rental expenses this period" value={fmtEur(rentalExp.total)} sub="Included in operating expenses" />
          <StatCard label="Gross profit this period" value={fmtEur(totalProfit)} accent={totalProfit < 0 ? C.danger : C.accent} sub="Revenue − labor cost" />
          <StatCard label="Gross margin" value={overallMargin !== null ? overallMargin.toFixed(0) + "%" : "—"} accent={overallMargin !== null && overallMargin < 0 ? C.danger : C.accent} sub="(Revenue − cost) ÷ revenue" />
          <StatCard label="Operating expenses" value={fmtEur(opex.total)} sub="Insurance, software, bank, depreciation, AI, business dev." />
          <StatCard label="Operating profit" value={fmtEur(operatingProfit)} accent={operatingProfit < 0 ? C.danger : C.accent} sub="Gross profit − operating expenses" />
          <StatCard label="Operating margin" value={operatingMargin !== null ? operatingMargin.toFixed(0) + "%" : "—"} accent={operatingMargin !== null && operatingMargin < 0 ? C.danger : C.accent} sub="Operating profit ÷ revenue" />
          <StatCard label="Avg. effective rate" value={fmtEur(avgEffectiveRate) + "/h"} sub="Revenue ÷ hours worked" />
        </div>
      )}

      {config.billingPanel && (() => {
        const lines = computeBillingLines(clients, billingStatuses);
        const overdue = lines.filter((l) => l.status === "overdue");
        const unpaid = lines.filter((l) => ["invoiced", "partial", "overdue"].includes(l.status));
        const notInvoiced = lines.filter((l) => l.status === "pending");
        if (unpaid.length === 0 && notInvoiced.length === 0) return null;
        const sum = (arr, f) => arr.reduce((t, l) => t + f(l), 0);
        const overdueAmt = sum(overdue, (l) => l.outstanding), unpaidAmt = sum(unpaid, (l) => l.outstanding), pendingAmt = sum(notInvoiced, (l) => l.amount);
        const cname = (id) => clients.find((c) => c.id === id)?.name || "—";
        // Worst first: overdue by days, then other unpaid by days, then not invoiced by month.
        const worst = [...overdue].sort((a, b) => (b.daysSinceInvoice || 0) - (a.daysSinceInvoice || 0));
        return (
          <Panel style={{ marginBottom: 18, borderLeft: `4px solid ${overdue.length ? C.danger : C.warn}` }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, gap: 10, flexWrap: "wrap" }}>
              <div>
                <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink }}>Billing — {overdue.length ? `${fmtEur(overdueAmt)} overdue` : `${fmtEur(unpaidAmt)} awaiting payment`}</div>
                <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted }}>
                  {overdue.length} overdue line{overdue.length === 1 ? "" : "s"} (invoiced more than {BILLING_OVERDUE_DAYS} days ago) · {fmtEur(unpaidAmt)} invoiced and unpaid · {fmtEur(pendingAmt)} not invoiced yet ({notInvoiced.length} line{notInvoiced.length === 1 ? "" : "s"})
                </div>
              </div>
              <Btn size="sm" variant="secondary" onClick={onOpenBilling}>Open billing</Btn>
            </div>
            {worst.slice(0, 5).map((l) => (
              <div key={l.key} style={{ display: "flex", gap: 10, padding: "5px 0", borderBottom: `1px solid ${C.border}`, fontFamily: sans, fontSize: 12.5, alignItems: "center" }}>
                <Badge tone="danger">Overdue</Badge>
                <span style={{ fontWeight: 600, flex: 1 }}>{cname(l.clientId)} <span style={{ fontWeight: 400, color: C.inkMuted }}>· {l.kind === "fixed" ? "Fixed fee" : `Extra${l.note ? ` — ${l.note}` : ""}`} {l.month}</span></span>
                <span style={{ fontFamily: mono, color: C.inkMuted }}>{l.daysSinceInvoice}d</span>
                <span style={{ fontFamily: mono, fontWeight: 700, minWidth: 90, textAlign: "right" }}>{fmtEur(l.outstanding)}</span>
              </div>
            ))}
            {worst.length > 5 && <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, paddingTop: 6 }}>…and {worst.length - 5} more overdue line{worst.length - 5 === 1 ? "" : "s"}.</div>}
          </Panel>
        );
      })()}

      {config.ddPanel && (() => {
        // Δέουσα επιμέλεια per client: coverage of the client base and the
        // client-level obligations that need action.
        const DD_KINDS = ["file", "approval", "review", "kmpd"];
        const ddAlerts = complianceAlerts.filter((a) => a.clientId && DD_KINDS.includes(a.kind));
        const active = clients.filter((c) => ["live", "onboarding"].includes(c.lifecycle));
        if (active.length === 0) return null;
        const rec = (c) => amlRecords.find((a) => a.clientId === c.id) || null;
        const completed = active.filter((c) => rec(c)?.status === "completed").length;
        const inProgress = active.filter((c) => rec(c) && ["in_progress"].includes(rec(c).status)).length;
        const noFile = active.filter((c) => !rec(c) || rec(c).status === "not_started").length;
        const highAwaiting = active.filter((c) => rec(c)?.riskLevel === "high" && rec(c).status !== "declined" && !rec(c).approvedBy).length;
        const reviewsOverdue = active.filter((c) => rec(c)?.nextReviewOn && daysUntil(rec(c).nextReviewOn) < 0).length;
        const reviewsSoon = active.filter((c) => rec(c)?.nextReviewOn && daysUntil(rec(c).nextReviewOn) >= 0 && daysUntil(rec(c).nextReviewOn) <= 30).length;
        const kmpd = ddAlerts.filter((a) => a.kind === "kmpd").length;
        if (ddAlerts.length === 0 && completed === active.length) return null;
        const urgent = ddAlerts.filter((a) => a.severity === "high").length;
        const pct = Math.round((completed / active.length) * 100);
        const items = [
          { label: "Χωρίς φάκελο", value: noFile, bad: noFile > 0 },
          { label: "Σε εξέλιξη", value: inProgress, bad: false, warn: inProgress > 0 },
          { label: "Υψηλός κίνδυνος χωρίς έγκριση", value: highAwaiting, bad: highAwaiting > 0 },
          { label: "Επανεξέταση σε καθυστέρηση", value: reviewsOverdue, bad: reviewsOverdue > 0 },
          { label: "Επανεξέταση ≤ 30 ημέρες", value: reviewsSoon, bad: false, warn: reviewsSoon > 0 },
          { label: "ΚΜΠΔ — μεταβολή UBO (60 ημ.)", value: kmpd, bad: kmpd > 0 },
        ];
        return (
          <Panel style={{ marginBottom: 18, borderLeft: `4px solid ${urgent > 0 ? C.danger : C.warn}` }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, gap: 10, flexWrap: "wrap" }}>
              <div>
                <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink }}>Δέουσα επιμέλεια πελατών — {completed}/{active.length} φάκελοι ολοκληρωμένοι{urgent > 0 ? ` · ${urgent} επείγοντα` : ""}</div>
                <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted }}>ν. 4557/2018 — ταυτοποίηση πριν την έναρξη της σχέσης (άρ. 14), έγκριση υψηλού κινδύνου (άρ. 16Α, 18), επανεξέταση (άρ. 13 παρ. 7), ΚΜΠΔ (άρ. 20).</div>
              </div>
              <Btn size="sm" variant="secondary" onClick={onOpenCompliance}>Open compliance</Btn>
            </div>
            <div style={{ height: 8, background: C.border, borderRadius: 4, overflow: "hidden", marginBottom: 10 }}>
              <div style={{ width: `${pct}%`, height: "100%", background: pct === 100 ? C.accent : C.warn }} />
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
              {items.map((it) => (
                <div key={it.label} style={{ border: `1px solid ${it.bad ? C.danger : it.warn ? C.warn : C.border}`, background: it.bad ? "#FDECEC" : "transparent", borderRadius: 7, padding: "5px 10px", fontFamily: sans, fontSize: 12 }}>
                  <span style={{ fontFamily: mono, fontWeight: 700, fontSize: 14, color: it.bad ? C.danger : it.warn ? C.warn : C.inkFaint, marginRight: 6 }}>{it.value}</span>{it.label}
                </div>
              ))}
            </div>
            {ddAlerts.slice(0, 8).map((a, i) => (
              <div key={i} style={{ display: "flex", gap: 10, padding: "5px 0", borderBottom: `1px solid ${C.border}`, fontFamily: sans, fontSize: 12.5, alignItems: "baseline" }}>
                <Badge tone={a.severity === "high" ? "danger" : "warn"}>{a.severity === "high" ? "Urgent" : "Upcoming"}</Badge>
                <span style={{ flex: 1 }}>
                  <button onClick={() => onOpenCard(a.clientId, "dd")} style={{ ...iconBtnStyle, padding: 0, fontFamily: sans, fontSize: 12.5, fontWeight: 600, color: C.accentDark, textDecoration: "underline" }}>{a.title}</button>
                  <span style={{ color: C.inkMuted }}> — {a.detail}</span>
                </span>
                <span style={{ fontFamily: mono, color: C.inkMuted }}>{a.due || ""}</span>
              </div>
            ))}
            {ddAlerts.length > 8 && <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, paddingTop: 6 }}>…και {ddAlerts.length - 8} ακόμα — δες το tab Alerts στο AML Compliance.</div>}
          </Panel>
        );
      })()}

      {config.compliancePanel && (() => {
        const rest = complianceAlerts.filter((a) => !(a.clientId && ["file", "approval", "review", "kmpd"].includes(a.kind)));
        if (rest.length === 0) return null;
        return (
          <Panel style={{ marginBottom: 18, borderLeft: `4px solid ${rest.some((a) => a.severity === "high") ? C.danger : C.warn}` }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <div>
                <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink }}>AML compliance γραφείου — {rest.filter((a) => a.severity === "high").length} urgent, {rest.filter((a) => a.severity !== "high").length} upcoming</div>
                <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted }}>Υποχρεώσεις του γραφείου, αναφορές προς την Αρχή και αιτήματα αρχών (ΑΑΔΕ, Αρχή) — ν. 4557/2018.</div>
              </div>
              <Btn size="sm" variant="secondary" onClick={onOpenCompliance}>Open compliance</Btn>
            </div>
            {rest.slice(0, 6).map((a, i) => (
              <div key={i} style={{ display: "flex", gap: 10, padding: "5px 0", borderBottom: `1px solid ${C.border}`, fontFamily: sans, fontSize: 12.5 }}>
                <Badge tone={a.severity === "high" ? "danger" : "warn"}>{a.severity === "high" ? "Urgent" : "Upcoming"}</Badge>
                <span style={{ fontWeight: 600, flex: 1 }}>{a.clientId ? <button onClick={() => onOpenCard(a.clientId, "dd")} style={{ ...iconBtnStyle, padding: 0, fontFamily: sans, fontSize: 12.5, fontWeight: 600, color: C.accentDark, textDecoration: "underline" }}>{a.title}</button> : a.title}</span>
                <span style={{ fontFamily: mono, color: C.inkMuted }}>{a.due || ""}</span>
              </div>
            ))}
          </Panel>
        );
      })()}

      {config.onboardingPanel && (() => {
        const leads = clients.filter((c) => c.lifecycle === "lead");
        const onboarding = clients.filter((c) => c.lifecycle === "onboarding");
        const recentLive = clients.filter((c) => c.lifecycle === "live" && c.goLiveDate && fromKey(c.goLiveDate) >= addDays(TODAY, -30));
        if (leads.length === 0 && onboarding.length === 0 && recentLive.length === 0) return null;
        const empName = (id) => employees.find((e) => e.id === id)?.name || "—";
        const link = (c) => <button onClick={() => onOpenCard && onOpenCard(c.id)} style={{ ...iconBtnStyle, padding: 0, fontFamily: sans, fontSize: 13, color: C.accentDark, textDecoration: "underline" }}>{c.name}</button>;
        return (
          <Panel style={{ marginBottom: 18 }}>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Pipeline</div>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>{leads.length} lead{leads.length === 1 ? "" : "s"} · {onboarding.length} in onboarding · {recentLive.length} went live in the last 30 days</div>
            <TableShell headers={["Client", "Stage", "Account manager", "Progress", "Next due"]}>
              {[...onboarding, ...leads, ...recentLive].map((c) => {
                const p = onboardingProgress(c, tasks);
                const next = p.list.filter((t) => t.status !== "done" && t.dueDate).sort((a, b) => a.dueDate.localeCompare(b.dueDate))[0];
                return (
                  <tr key={c.id}>
                    <Td>{link(c)}</Td>
                    <Td><Badge tone={LIFECYCLE_TONE[c.lifecycle]}>{LIFECYCLE_LABELS[c.lifecycle]}</Badge>{c.lifecycle === "live" && c.goLiveDate ? <span style={{ fontFamily: mono, fontSize: 11, color: C.inkMuted, marginLeft: 6 }}>{c.goLiveDate}</span> : null}</Td>
                    <Td>{empName(c.accountManagerId)}</Td>
                    <Td>
                      {c.lifecycle === "onboarding" && p.total > 0 ? (
                        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <div style={{ width: 120, height: 7, background: C.border, borderRadius: 4, overflow: "hidden" }}><div style={{ width: `${(p.done / p.total) * 100}%`, height: "100%", background: C.warn }} /></div>
                          <span style={{ fontFamily: mono, fontSize: 11.5, color: C.inkMuted }}>{p.done}/{p.total}</span>
                        </div>
                      ) : c.lifecycle === "onboarding" ? <span style={{ color: C.inkFaint }}>no checklist</span> : "—"}
                    </Td>
                    <Td mono style={{ color: next && taskIsOverdue(next) ? C.danger : C.inkMuted }}>{next ? `${next.dueDate} · ${next.title}` : "—"}</Td>
                  </tr>
                );
              })}
            </TableShell>
          </Panel>
        );
      })()}

      {config.tasksPanel && (() => {
        const weekEnd = addDays(TODAY, 7);
        const open = tasks.filter((t) => t.status !== "done");
        const overdue = open.filter(taskIsOverdue).sort((a, b) => a.dueDate.localeCompare(b.dueDate));
        const dueSoon = open.filter((t) => t.dueDate && !taskIsOverdue(t) && fromKey(t.dueDate) <= weekEnd).sort((a, b) => a.dueDate.localeCompare(b.dueDate));
        const show = [...overdue, ...dueSoon];
        const empName = (id) => employees.find((e) => e.id === id)?.name || "Unassigned";
        const clientLabel = (id) => clients.find((c) => c.id === id)?.name || "—";
        return (
          <Panel style={{ marginBottom: 18 }}>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Tasks</div>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>
              {overdue.length} overdue · {dueSoon.length} due in the next 7 days · {open.length} open in total
            </div>
            {show.length === 0 ? (
              <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted }}>Nothing overdue or due this week.</div>
            ) : (
              <TableShell headers={["Due", "Task", "Client", "Assigned to", "Priority", "Status"]}>
                {show.slice(0, 15).map((t) => (
                  <tr key={t.id}>
                    <Td mono style={{ color: taskIsOverdue(t) ? C.danger : undefined }}>{t.dueDate}</Td>
                    <Td>{t.title}</Td>
                    <Td>{t.clientId ? <button onClick={() => onOpenCard && onOpenCard(t.clientId)} style={{ ...iconBtnStyle, padding: 0, fontFamily: sans, fontSize: 13, color: C.accentDark, textDecoration: "underline" }}>{clientLabel(t.clientId)}</button> : "—"}</Td>
                    <Td>{empName(t.assignedTo)}</Td>
                    <Td><Badge tone={TASK_PRIORITY_TONE[t.priority]}>{TASK_PRIORITIES[t.priority]}</Badge></Td>
                    <Td><Badge tone={taskIsOverdue(t) ? "danger" : TASK_STATUS_TONE[t.status]}>{taskIsOverdue(t) ? "Overdue" : TASK_STATUSES[t.status]}</Badge></Td>
                  </tr>
                ))}
              </TableShell>
            )}
          </Panel>
        );
      })()}

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
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 4 }}>Hours by client</div>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Effective rate = revenue of the period ÷ hours worked for the client. Compare it with cost per hour to see whether the fee covers the time spent.</div>
            <TableShell headers={["Client", "Hours", "Allocation", "Status", "Revenue", "Effective rate", "Cost", "Cost /h", "Profit", "Margin"]}>
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
                    <Td mono>{c.hours > 0 ? fmtEur(c.revenue / c.hours) + "/h" : "—"}</Td>
                    <Td mono>{fmtEur(c.cost)}</Td>
                    <Td mono>{c.hours > 0 ? fmtEur(c.cost / c.hours) + "/h" : "—"}</Td>
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
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 13.5, color: C.ink, margin: "18px 0 4px" }}>Operating result</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 10 }}>Gross profit − operating expenses = operating profit. Expenses booked as "General" are company-wide and not attributed to a department.</div>
          <TableShell headers={["", "Gross profit", "− Operating expenses", "= Operating profit", "Operating margin"]}>
            {[
              ...byDepartment.map((d) => ({ key: d.key, label: d.label, gross: d.profit, ox: opex.byDepartment[d.key] || 0, revenue: d.revenue })),
              ...(Math.abs(opex.byDepartment[OPEX_GENERAL]) > 0.005 || (revenueByCC[UNALLOCATED] || 0) > 0.005 ? [{ key: OPEX_GENERAL, label: "General / not allocated", gross: revenueByCC[UNALLOCATED] || 0, ox: opex.byDepartment[OPEX_GENERAL] || 0, revenue: revenueByCC[UNALLOCATED] || 0 }] : []),
            ].concat([{ key: "total", label: "Company total", gross: totalProfit, ox: opex.total, revenue: totalRevenue }]).map((r) => {
              const op = r.gross - r.ox; const mg = r.revenue > 0 ? (op / r.revenue) * 100 : null;
              return (
                <tr key={r.key} style={{ fontWeight: r.key === "total" ? 700 : 400, borderTop: r.key === "total" ? `2px solid ${C.borderStrong}` : undefined }}>
                  <Td>{r.label}</Td>
                  <Td mono>{fmtEur(r.gross)}</Td>
                  <Td mono>{fmtEur(r.ox)}</Td>
                  <Td mono style={{ color: op < 0 ? C.danger : C.accentDark, fontWeight: 700 }}>{fmtEur(op)}</Td>
                  <Td>{mg !== null ? <Badge tone={mg < 0 ? "danger" : mg < 20 ? "warn" : "accent"}>{mg.toFixed(1)}%</Badge> : "—"}</Td>
                </tr>
              );
            })}
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
            Fixed fees follow each client's cost-center split, extra fees their own tag. Labor cost follows each employee's department: Accounting Team by logged hours, other departments in proportion to the department's revenue by cost center. Rental expenses (rent, utilities, common expenses) are operating expenses, shown below gross profit.
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
          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>New clients, revenue, labor cost, and gross profit — last 6 months.</div>
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

        {config.clientProfitTable && (
          <Panel style={{ flex: "1 1 100%", minWidth: 380 }}>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Profitability by client — gross and operating margin</div>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>
              Gross margin = (revenue − labor cost of the hours on the client) ÷ revenue. Operating margin also charges each client its share of the overhead (cost of internal hours + operating expenses), allocated by cost center in proportion to the hours spent on the client — see the table below. Sorted from lowest operating margin.
            </div>
            <TableShell headers={["Client", "Hours", "Revenue", "Labor cost", "Gross profit", "Gross margin", "Overhead", "Operating profit", "Operating margin"]}>
              {clientProfitRows.map((r) => (
                <tr key={r.id}>
                  <Td>{r.name}</Td>
                  <Td mono>{r.hours.toFixed(1)}</Td>
                  <Td mono>{fmtEur(r.revenue)}</Td>
                  <Td mono>{fmtEur(r.cost)}</Td>
                  <Td mono>{fmtEur(r.grossProfit)}</Td>
                  <Td mono style={{ fontWeight: 700, color: r.grossMargin < 0 ? C.danger : undefined }}>{r.grossMargin.toFixed(0)}%</Td>
                  <Td mono>{fmtEur(r.overhead)}</Td>
                  <Td mono style={{ color: r.operatingProfit < 0 ? C.danger : undefined }}>{fmtEur(r.operatingProfit)}</Td>
                  <Td mono style={{ fontWeight: 700, color: r.operatingMargin < 0 ? C.danger : undefined }}>{r.operatingMargin.toFixed(0)}%</Td>
                </tr>
              ))}
              {clientProfitRows.length > 0 && (() => {
                const t = clientProfitRows.reduce((a, r) => ({ h: a.h + r.hours, rev: a.rev + r.revenue, c: a.c + r.cost, gp: a.gp + r.grossProfit, oh: a.oh + r.overhead, op: a.op + r.operatingProfit }), { h: 0, rev: 0, c: 0, gp: 0, oh: 0, op: 0 });
                return (
                  <tr style={{ background: C.accentSoft }}>
                    <Td style={{ fontWeight: 700 }}>Total</Td>
                    <Td mono style={{ fontWeight: 700 }}>{t.h.toFixed(1)}</Td>
                    <Td mono style={{ fontWeight: 700 }}>{fmtEur(t.rev)}</Td>
                    <Td mono style={{ fontWeight: 700 }}>{fmtEur(t.c)}</Td>
                    <Td mono style={{ fontWeight: 700 }}>{fmtEur(t.gp)}</Td>
                    <Td mono style={{ fontWeight: 700 }}>{t.rev > 0 ? ((t.gp / t.rev) * 100).toFixed(0) + "%" : "—"}</Td>
                    <Td mono style={{ fontWeight: 700 }}>{fmtEur(t.oh)}</Td>
                    <Td mono style={{ fontWeight: 700 }}>{fmtEur(t.op)}</Td>
                    <Td mono style={{ fontWeight: 700 }}>{t.rev > 0 ? ((t.op / t.rev) * 100).toFixed(0) + "%" : "—"}</Td>
                  </tr>
                );
              })()}
            </TableShell>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 13, color: C.ink, margin: "16px 0 2px" }}>Overhead by cost center</div>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 8 }}>Pool = labor cost not charged to clients + operating expenses of the cost center. Driver = client hours (Accounting / Tax / Payroll / Other: hours logged with that category; Director, Finance, Rental, General: total client hours).</div>
            <TableShell headers={["Cost center", "Overhead pool", "Driver hours", "Overhead per hour"]}>
              {overheadByCC.map((r) => (
                <tr key={r.key}>
                  <Td>{r.label}</Td>
                  <Td mono>{fmtEur(r.pool)}</Td>
                  <Td mono>{r.hours.toFixed(1)}</Td>
                  <Td mono>{r.perHour != null ? "€" + r.perHour.toFixed(0) + "/h" : "split equally"}</Td>
                </tr>
              ))}
              <tr style={{ background: C.accentSoft }}>
                <Td style={{ fontWeight: 700 }}>Total</Td>
                <Td mono style={{ fontWeight: 700 }}>{fmtEur(overheadByCC.reduce((s, r) => s + r.pool, 0))}</Td>
                <Td /><Td />
              </tr>
            </TableShell>
          </Panel>
        )}

        {config.effectiveRateTable && (
          <Panel style={{ flex: 1, minWidth: 380 }}>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Effective rate by client, by month</div>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Revenue ÷ hours, €/h. Red = below the cost per hour. Sorted from lowest to highest 6-month rate.</div>
            <TableShell headers={["Client", ...trendMonths.map((m) => m.label), "6-month rate", "Cost /h"]}>
              {effRateRows.map((r) => (
                <tr key={r.id}>
                  <Td>{r.name}</Td>
                  {r.cells.map((x, i) => <Td key={i} mono style={{ color: x.rate != null && x.costH != null && x.rate < x.costH ? C.danger : undefined }}>{x.rate != null ? "€" + x.rate.toFixed(0) : "—"}</Td>)}
                  <Td mono style={{ fontWeight: 700, color: r.rate != null && r.costH != null && r.rate < r.costH ? C.danger : undefined }}>{r.rate != null ? "€" + r.rate.toFixed(0) + "/h" : "—"}</Td>
                  <Td mono>{r.costH != null ? "€" + r.costH.toFixed(0) + "/h" : "—"}</Td>
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

function AdminTimesheets({ employees, clients, entries, refetchEntries, lockedWeeks, refetchLockedWeeks, leaveRequests = [] }) {
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
  // Approved leave on a working day (not weekend, not a public holiday) lowers that week's contracted hours.
  const weekHolidays = holidayMapForYears([weekStart.getFullYear(), weekEnd.getFullYear()]);
  const weekDays = Array.from({ length: 7 }, (_, i) => addDays(weekStart, i)).filter((d) => !isWeekend(d) && !weekHolidays.has(toKey(d)));
  const compliance = employees.filter((e) => employeeBilledInRange(e, weekStart, weekEnd)).map((e) => {
    const base = e.weeklyHours || 40;
    const approved = leaveRequests.filter((r) => r.employeeId === e.id && r.status === "approved");
    const leaveDays = weekDays.filter((d) => approved.some((r) => r.startDate <= toKey(d) && toKey(d) <= r.endDate)).length;
    const target = Math.max(0, Math.round(base * (1 - leaveDays / 5) * 100) / 100);
    const logged = weekEntries.filter((t) => t.employeeId === e.id).reduce((s, t) => s + t.hours, 0);
    const diff = logged - target;
    return { ...e, base, leaveDays, target, logged, diff };
  }).sort((a, b) => b.diff - a.diff);

  return (
    <div>
      <PageHeader title="All Timesheets" sub="Review and correct any employee's logged hours." />

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 4 }}>
          Weekly hours vs. contract
        </div>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>
          {fmtShort(weekStart)} – {fmtShort(weekEnd)}. Contracted hours are set on the Employees page; approved leave days reduce the week's target.
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
                <Td mono>{e.target}h{e.leaveDays > 0 && <span style={{ display: "block", fontFamily: sans, fontSize: 11, color: C.inkMuted }}>{e.base}h − {e.leaveDays} leave day{e.leaveDays > 1 ? "s" : ""}</span>}</Td>
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

function AdminClients({ user, clients, refetchClients, entries, onOpenCard, tasks = [], templates = [], refetchTasks, amlRecords = [], employees = [] }) {
  const [showAmlRegister, setShowAmlRegister] = useState(false);
  const [modal, setModal] = useState(null); // null | "new" | client object
  const [name, setName] = useState("");
  const [lifecycleNew, setLifecycleNew] = useState("onboarding");
  const [newCompany, setNewCompany] = useState(false);
  const [lifecycleFilter, setLifecycleFilter] = useState("all");
  const [hideInactive, setHideInactive] = useState(() => { try { return localStorage.getItem("gloria_hide_inactive") === "1"; } catch { return false; } });
  const toggleHideInactive = (v) => { setHideInactive(v); try { localStorage.setItem("gloria_hide_inactive", v ? "1" : "0"); } catch { /* ignore */ } };
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
    setModal("new"); setName(""); setFixedFee(""); setFeeEffectiveDate(toKey(TODAY)); setEndDate(""); setEditingHistoryId(null); setLifecycleNew("onboarding"); setNewCompany(false);
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
      const { data, error } = await supabase.from("clients").insert({ name: name.trim(), fixed_fee: fee, end_date: endDate || null, lifecycle: lifecycleNew, go_live_date: lifecycleNew === "live" ? toKey(TODAY) : null }).select().single();
      if (error || !data) {
        window.alert(`Couldn't create client: ${error ? error.message : "unknown error"}`);
        return;
      }
      if (lifecycleNew === "onboarding") {
        await startOnboarding({ id: data.id, accountManagerId: null, legalForm: "" }, templates, [], user, { includeFormation: newCompany });
        if (refetchTasks) refetchTasks();
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
  const filtered = clients
    .filter((c) => c.name.toLowerCase().includes(query.toLowerCase()))
    .filter((c) => lifecycleFilter === "all" || c.lifecycle === lifecycleFilter)
    .filter((c) => !hideInactive || c.active);
  const inactiveCount = clients.filter((c) => !c.active).length;
  const liveClient = modal && modal !== "new" ? clients.find((c) => c.id === modal.id) : null;
  const lifecycleCounts = Object.fromEntries(Object.keys(LIFECYCLE_LABELS).map((k) => [k, clients.filter((c) => c.lifecycle === k).length]));

  return (
    <div>
      <PageHeader title="Clients" sub="Manage the clients your team logs hours against, and their billing terms."
        right={
          <div style={{ display: "flex", gap: 8 }}>
            <Btn variant="secondary" onClick={() => exportToExcel(`clients-${toKey(TODAY)}.xlsx`, [{
              name: "Clients",
              rows: clients.map((c) => ({
                Name: c.name, Stage: LIFECYCLE_LABELS[c.lifecycle], "Go-live date": c.goLiveDate || "", Status: c.active ? "Active" : "Inactive",
                "Fixed fee (€/mo)": c.fixedFee ?? "", "Accounting alloc. (h)": c.allocation?.accounting || 0,
                "Tax alloc. (h)": c.allocation?.tax || 0, "Payroll alloc. (h)": c.allocation?.payroll || 0,
                "Other alloc. (h)": c.allocation?.other || 0,
                ...Object.fromEntries(COST_CENTERS.map((cc) => [`Fee → ${COST_CENTER_LABELS[cc]} (€)`, c.feeSplit?.[cc] || 0])),
                "Fee → Other note": c.feeSplitNote || "",
              })),
            }])}>Export to Excel</Btn>
            {user.role === "admin" && <Btn variant="secondary" onClick={() => setShowAmlRegister((v) => !v)}>{showAmlRegister ? "Hide AML register" : "AML register"}</Btn>}
            <Btn icon={Plus} onClick={openNew}>Add client</Btn>
          </div>
        } />
      {showAmlRegister && (() => {
        // Συγκεντρωτική κατάσταση κατάταξης πελατών — one line per client,
        // the list an ΑΑΔΕ audit asks for (question 15 of the questionnaire).
        const empName = (id) => employees.find((e) => e.id === id)?.name || "—";
        const rows = clients.filter((c) => c.lifecycle !== "lead").map((c) => ({ c, a: amlRecords.find((x) => x.clientId === c.id) || null }))
          .sort((x, y) => (x.a ? 0 : 1) - (y.a ? 0 : 1) || x.c.name.localeCompare(y.c.name));
        const counts = { assessed: rows.filter((r) => r.a && r.a.status === "completed").length, inProgress: rows.filter((r) => r.a && r.a.status !== "completed" && r.a.status !== "declined").length, none: rows.filter((r) => !r.a).length, review: rows.filter((r) => r.a?.nextReviewOn && fromKey(r.a.nextReviewOn) < TODAY).length };
        return (
          <Panel style={{ marginBottom: 18 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
              <div>
                <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink }}>AML register — client risk classification (ν. 4557/2018)</div>
                <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted }}>{counts.assessed} assessed · {counts.inProgress} in progress · {counts.none} without a file · {counts.review} review overdue</div>
              </div>
              <Btn size="sm" variant="secondary" onClick={() => exportToExcel(`aml-register-${toKey(TODAY)}.xlsx`, [{ name: "AML register", rows: rows.map(({ c, a }) => ({
                Client: c.name, "ΑΦΜ": c.afm || "", "Legal form": c.legalForm || "", Stage: LIFECYCLE_LABELS[c.lifecycle], "Entity type (ΠΟΛ.1200)": a ? (AML_ENTITY_TYPES[a.entityType] || "") : "",
                "Risk sum": a ? a.riskScore : "", "Risk level": a ? AML_LEVEL_LABELS[a.riskLevel] : "", "Due diligence": a ? AML_DD_LABELS[a.ddLevel] : "", PEP: a ? (a.pepStatus !== "none" ? "Yes" : "No") : "",
                "File status": a ? AML_STATUS_LABELS[a.status] : "No file", "Completed on": a?.completedOn || "", "Completed by": a ? empName(a.completedBy) : "", "Approved by": a ? empName(a.approvedBy) : "", "Next review": a?.nextReviewOn || "",
              })) }])}>Export to Excel</Btn>
            </div>
            <TableShell headers={["Client", "Stage", "Risk", "Due diligence", "PEP", "File", "Completed", "Next review", ""]}>
              {rows.map(({ c, a }) => (
                <tr key={c.id}>
                  <Td style={{ fontWeight: 600 }}>{c.name}</Td>
                  <Td><Badge tone={LIFECYCLE_TONE[c.lifecycle]}>{LIFECYCLE_LABELS[c.lifecycle]}</Badge></Td>
                  <Td>{a ? <><Badge tone={AML_LEVEL_TONE[a.riskLevel]}>{AML_LEVEL_LABELS[a.riskLevel]}</Badge><span style={{ fontFamily: mono, fontSize: 11, color: C.inkMuted, marginLeft: 6 }}>{a.riskScore >= 0 ? "+" : ""}{a.riskScore}</span></> : <span style={{ color: C.inkFaint }}>—</span>}</Td>
                  <Td style={{ color: C.inkMuted }}>{a ? AML_DD_LABELS[a.ddLevel] : "—"}</Td>
                  <Td>{a && a.pepStatus !== "none" ? <Badge tone="danger">PEP</Badge> : <span style={{ color: C.inkFaint }}>—</span>}</Td>
                  <Td>{a ? <Badge tone={AML_STATUS_TONE[a.status]}>{AML_STATUS_LABELS[a.status]}</Badge> : <Badge tone="danger">No file</Badge>}</Td>
                  <Td mono style={{ color: C.inkMuted }}>{a?.completedOn || "—"}</Td>
                  <Td mono style={{ color: a?.nextReviewOn && fromKey(a.nextReviewOn) < TODAY ? C.danger : C.inkMuted }}>{a?.nextReviewOn || "—"}</Td>
                  <Td right><Btn size="sm" variant="ghost" onClick={() => onOpenCard(c.id)}>Open</Btn></Td>
                </tr>
              ))}
            </TableShell>
          </Panel>
        );
      })()}
      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 10 }}>
          Hours, allocation, and fee below reflect this reporting period — not just the latest settings.
        </div>
        <PeriodSelector period={viewPeriod} setPeriod={setViewPeriod} anchor={viewAnchor} setAnchor={setViewAnchor} />
      </Panel>
      <Panel style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 14, flexWrap: "wrap" }}>
          <div style={{ position: "relative", width: 260 }}>
            <Search size={14} color={C.inkFaint} style={{ position: "absolute", left: 9, top: 9 }} />
            <input placeholder="Search clients…" value={query} onChange={(e) => setQuery(e.target.value)}
              style={{ ...inputStyle, width: "100%", paddingLeft: 28 }} />
          </div>
          <div style={{ display: "flex", border: `1px solid ${C.border}`, borderRadius: 7, overflow: "hidden" }}>
            {[["all", "All"], ...Object.entries(LIFECYCLE_LABELS)].map(([k, label]) => (
              <button key={k} type="button" onClick={() => setLifecycleFilter(k)}
                style={{ padding: "7px 11px", fontSize: 12.5, fontFamily: sans, fontWeight: 600, border: "none", cursor: "pointer",
                  background: lifecycleFilter === k ? C.accent : C.surface, color: lifecycleFilter === k ? "#fff" : C.inkMuted }}>
                {label}{k !== "all" ? ` (${lifecycleCounts[k]})` : ""}
              </button>
            ))}
          </div>
          <label style={{ display: "flex", alignItems: "center", gap: 6, fontFamily: sans, fontSize: 12.5, color: C.ink, cursor: "pointer" }}>
            <input type="checkbox" checked={hideInactive} onChange={(e) => toggleHideInactive(e.target.checked)} /> Hide deactivated ({inactiveCount})
          </label>
        </div>
        <TableShell headers={["Client", "Stage", "Hours", "Revenue", "Allocation", "Fixed fee", "Extra fees (period)", "Status", ""]}>
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
                <Td>
                  <Badge tone={LIFECYCLE_TONE[c.lifecycle]}>{LIFECYCLE_LABELS[c.lifecycle]}</Badge>
                  {c.lifecycle === "onboarding" && (() => { const p = onboardingProgress(c, tasks); return <span style={{ fontFamily: mono, fontSize: 11, color: C.inkMuted, marginLeft: 6 }}>{p.done}/{p.total}</span>; })()}
                </Td>
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
                    <Btn size="sm" variant="secondary" icon={UserRound} onClick={() => onOpenCard(c.id)}>Card</Btn>
                    <Btn size="sm" variant="ghost" icon={Euro} onClick={() => onOpenCard(c.id, "fees")}>Fees</Btn>
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
                <Td />
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
            {modal === "new" && (
              <Field label="Stage">
                <select style={inputStyle} value={lifecycleNew} onChange={(e) => setLifecycleNew(e.target.value)}>
                  <option value="lead">Lead — prospect, no setup yet</option>
                  <option value="onboarding">Onboarding — create the setup checklist now ({templates.filter((t) => t.active).length} tasks)</option>
                  <option value="live">Live — already set up</option>
                </select>
                {lifecycleNew === "onboarding" && (
                  <label style={{ display: "flex", alignItems: "center", gap: 8, fontFamily: sans, fontSize: 12.5, color: C.ink, marginTop: 8 }}>
                    <input type="checkbox" checked={newCompany} onChange={(e) => setNewCompany(e.target.checked)} />
                    Newly formed company — also create the company formation tasks (ΥΜΣ, μισθωτήριο, τράπεζα, ΕΦΚΑ, κεφάλαιο, ΓΕΜΗ…)
                  </label>
                )}
              </Field>
            )}
            {liveClient && user.role === "admin" && (() => {
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
    if (!window.confirm(`Permanently delete "${e.name}"? This removes their GLORIA record but does not revoke their login — this cannot be undone.`)) return;
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
              Accounting Team cost is spread across Accounting / Tax / Payroll / Other by the hours logged each period. Management Team cost is split between Director and Finance in proportion to that period's Director/Finance fees. Rental expenses are operating expenses, not part of department labor cost.
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
        <StatCard label="Total rental expenses" value={fmtEur(accrued.total)} accent={C.accent} sub="Counted in operating expenses" />
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


function AdminOpex({ expenses, refetchOpex, rentalExpenses = [] }) {
  const [viewPeriod, setViewPeriod] = useState("month");
  const [viewAnchor, setViewAnchor] = useState(TODAY);
  const [modal, setModal] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [kind, setKind] = useState("recurring");
  const [month, setMonth] = useState(toKey(TODAY).slice(0, 7));
  const [fromMonth, setFromMonth] = useState(toKey(TODAY).slice(0, 7));
  const [toMonth, setToMonth] = useState("");
  const [category, setCategory] = useState("insurance");
  const [department, setDepartment] = useState(OPEX_GENERAL);
  const [amount, setAmount] = useState("");
  const [note, setNote] = useState("");
  const [catFilter, setCatFilter] = useState("all");

  const [rangeStart, rangeEnd] = periodRange(viewPeriod, viewAnchor);
  const accrued = opexForRange(expenses, rangeStart, rangeEnd, rentalExpenses);
  const monthsInRange = [];
  for (let d = startOfMonth(rangeStart); d <= rangeEnd; d = new Date(d.getFullYear(), d.getMonth() + 1, 1)) monthsInRange.push(monthKeyOf(d));
  const startKeyOf = (x) => x.kind === "recurring" ? (x.fromMonth || "") : (x.month || "");
  const periodEntries = expenses
    .filter((x) => monthsInRange.some((m) => rentalExpenseAppliesTo(x, m)))
    .filter((x) => catFilter === "all" || x.category === catFilter)
    .sort((a, b) => (a.kind === b.kind ? 0 : a.kind === "recurring" ? -1 : 1) || startKeyOf(b).localeCompare(startKeyOf(a)));
  const whenLabel = (x) => x.kind === "recurring" ? `${fmtMonthKey(x.fromMonth)} → ${x.toMonth ? fmtMonthKey(x.toMonth) : "open"}` : fmtMonthKey(x.month);

  function openNew(k = "recurring") {
    setEditingId(null); setKind(k);
    const m = monthKeyOf(rangeStart);
    setMonth(m); setFromMonth(m); setToMonth(""); setCategory("accounting_software"); setDepartment(OPEX_GENERAL); setAmount(""); setNote("");
    setModal(true);
  }
  function openEdit(x) {
    setEditingId(x.id); setKind(x.kind);
    setMonth(x.month || monthKeyOf(rangeStart)); setFromMonth(x.fromMonth || monthKeyOf(rangeStart)); setToMonth(x.toMonth || "");
    setCategory(x.category); setDepartment(x.department); setAmount(String(x.amount)); setNote(x.note || "");
    setModal(true);
  }
  async function save() {
    if (kind === "recurring" && !(Number(amount) > 0)) { window.alert("Enter a monthly amount greater than zero."); return; }
    if (kind === "extra" && (amount === "" || Number.isNaN(Number(amount)) || Number(amount) === 0)) { window.alert("Enter an amount (negative for a credit or refund)."); return; }
    if (kind === "extra" && !month) { window.alert("Choose the month."); return; }
    if (kind === "recurring" && !fromMonth) { window.alert("Choose the month it starts."); return; }
    if (kind === "recurring" && toMonth && toMonth < fromMonth) { window.alert('"To" month must be the same as or after the "From" month.'); return; }
    if (category === "other" && !note.trim()) { window.alert('A note is required when the category is "Other".'); return; }
    const row = {
      kind, category, department, amount: Number(amount), note: note.trim() || null,
      month: kind === "extra" ? month : null, from_month: kind === "recurring" ? fromMonth : null, to_month: kind === "recurring" ? (toMonth || null) : null,
    };
    const { error } = editingId ? await supabase.from("operating_expenses").update(row).eq("id", editingId) : await supabase.from("operating_expenses").insert(row);
    if (error) { window.alert(`Couldn't save expense: ${error.message}`); return; }
    await refetchOpex();
    setModal(false);
  }
  async function remove(x) {
    if (!window.confirm(`Delete this ${OPEX_LABELS[x.category] || x.category} expense?`)) return;
    await supabase.from("operating_expenses").delete().eq("id", x.id);
    refetchOpex();
  }
  function exportOpex() {
    exportToExcel(`operating-expenses-${toKey(TODAY)}.xlsx`, [
      { name: "By category", rows: OPEX_DISPLAY_CATEGORIES.map((k) => ({ Category: OPEX_LABELS[k], "Amount (€)": Number(accrued.byCategory[k].toFixed(2)) })).concat([{ Category: "Total", "Amount (€)": Number(accrued.total.toFixed(2)) }]) },
      { name: "Entries", rows: periodEntries.map((x) => ({ Type: x.kind === "recurring" ? "Recurring" : "Extra", From: x.kind === "recurring" ? x.fromMonth : x.month, To: x.kind === "recurring" ? (x.toMonth || "open") : x.month, Category: OPEX_LABELS[x.category] || x.category, Department: opexDepartmentLabel(x.department), "Amount (€/month)": x.amount, Note: x.note })) },
    ]);
  }

  return (
    <div>
      <PageHeader title="Operating Expenses" sub="Other operating expenses — insurance, accounting software, bank charges, equipment depreciation, AI costs, business development. Gross profit minus these gives operating profit."
        right={<div style={{ display: "flex", gap: 8 }}>
          <Btn variant="secondary" onClick={exportOpex}>Export Excel</Btn>
          <Btn variant="secondary" icon={Plus} onClick={() => openNew("extra")}>Add extra</Btn>
          <Btn icon={Plus} onClick={() => openNew("recurring")}>Add recurring expense</Btn>
        </div>} />
      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 10 }}>
          Recurring expenses run from a start month to an end month or open-ended; extras are one-off amounts for a single month (negative = credit or refund). Everything is accrued to the reporting period, like rental expenses. Pick a department to attribute the cost, or "General" for company-wide costs.
        </div>
        <PeriodSelector period={viewPeriod} setPeriod={setViewPeriod} anchor={viewAnchor} setAnchor={setViewAnchor} />
      </Panel>
      <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
        {OPEX_DISPLAY_CATEGORIES.map((k) => <StatCard key={k} label={OPEX_LABELS[k].split(" (")[0]} value={fmtEur(accrued.byCategory[k])} sub={k === "rental" ? "From the Rental page" : "This period"} />)}
        <StatCard label="Total operating expenses" value={fmtEur(accrued.total)} accent={C.accent} sub="Deducted from gross profit" />
      </div>
      <Panel>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink }}>Expenses applying to this period</div>
          <select style={{ ...inputStyle, width: 280 }} value={catFilter} onChange={(e) => setCatFilter(e.target.value)}>
            <option value="all">All categories</option>
            {OPEX_CATEGORIES.map((k) => <option key={k} value={k}>{OPEX_LABELS[k]}</option>)}
          </select>
        </div>
        {periodEntries.length === 0 ? <EmptyState icon={Euro} title="No operating expenses booked for these months" sub="Add a recurring expense or an extra to start." /> : (
          <TableShell headers={["Type", "Period", "Category", "Department", "Amount / month", "Note", ""]}>
            {periodEntries.map((x) => (
              <tr key={x.id}>
                <Td><Badge tone={x.kind === "recurring" ? "accent" : "warn"}>{x.kind === "recurring" ? "Recurring" : "Extra"}</Badge></Td>
                <Td mono>{whenLabel(x)}</Td>
                <Td>{OPEX_LABELS[x.category] || x.category}</Td>
                <Td style={{ color: C.inkMuted }}>{opexDepartmentLabel(x.department)}</Td>
                <Td mono style={{ color: x.amount < 0 ? C.accentDark : undefined }}>{fmtEur(x.amount)}{x.amount < 0 ? " (credit)" : ""}</Td>
                <Td style={{ color: C.inkMuted }}>{x.note || "—"}</Td>
                <Td right>
                  <div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}>
                    <button onClick={() => openEdit(x)} style={iconBtnStyle} title="Edit"><Pencil size={13} color={C.inkMuted} /></button>
                    <button onClick={() => remove(x)} style={iconBtnStyle} title="Delete"><Trash2 size={13} color={C.inkMuted} /></button>
                  </div>
                </Td>
              </tr>
            ))}
          </TableShell>
        )}
      </Panel>
      {modal && (
        <Modal title={editingId ? "Edit operating expense" : "Add operating expense"} onClose={() => setModal(false)} width={440}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <div style={{ display: "flex", border: `1px solid ${C.border}`, borderRadius: 7, overflow: "hidden" }}>
              {Object.keys(RENTAL_EXPENSE_KINDS).map((k) => (
                <button key={k} type="button" onClick={() => setKind(k)} style={{ flex: 1, padding: "8px 10px", fontSize: 12.5, fontFamily: sans, fontWeight: 600, border: "none", cursor: "pointer", background: kind === k ? C.accent : C.surface, color: kind === k ? "#fff" : C.inkMuted }}>{RENTAL_EXPENSE_KINDS[k]}</button>
              ))}
            </div>
            <Field label="Category">
              <select style={inputStyle} value={category} onChange={(e) => setCategory(e.target.value)}>{OPEX_CATEGORIES.map((k) => <option key={k} value={k}>{OPEX_LABELS[k]}</option>)}</select>
            </Field>
            <Field label="Allocated to">
              <select style={inputStyle} value={department} onChange={(e) => setDepartment(e.target.value)}>
                <option value={OPEX_GENERAL}>{opexDepartmentLabel(OPEX_GENERAL)}</option>
                {DEPARTMENTS.map((d) => <option key={d.key} value={d.key}>{d.label}</option>)}
              </select>
            </Field>
            {kind === "extra" ? (
              <Field label="Month"><input type="month" style={inputStyle} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
            ) : (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <Field label="From month"><input type="month" style={inputStyle} value={fromMonth} onChange={(e) => setFromMonth(e.target.value)} /></Field>
                <Field label="To month (empty = open-ended)"><input type="month" style={inputStyle} value={toMonth} onChange={(e) => setToMonth(e.target.value)} /></Field>
              </div>
            )}
            <Field label={kind === "recurring" ? "Amount per month (€)" : "Amount (€) — negative for a credit or refund"}>
              <input type="number" min={kind === "recurring" ? "0" : undefined} style={inputStyle} value={amount} onChange={(e) => setAmount(e.target.value)} placeholder={kind === "recurring" ? "e.g. 85" : "e.g. 240 or -50"} autoFocus />
            </Field>
            <Field label={category === "other" ? "Note (required)" : "Note (optional)"}>
              <input style={inputStyle} value={note} onChange={(e) => setNote(e.target.value)} placeholder="e.g. Annual office insurance, split monthly" />
            </Field>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 6 }}>
              <Btn variant="secondary" onClick={() => setModal(false)}>Cancel</Btn>
              <Btn onClick={save}>{editingId ? "Save changes" : "Add expense"}</Btn>
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
    const totalCost = totalLaborCost; // rental expenses are operating expenses, below gross profit
    const revenueByCC = emptyByCC();
    byClient.forEach((c) => { Object.keys(revenueByCC).forEach((cc) => { revenueByCC[cc] += c.revenueByCostCenter?.[cc] || 0; }); });
    employeeCosts.forEach((ec) => {
      const emp = filteredEmployees.find((e) => e.id === ec.id);
      ec.byCostCenter = employeeCostByCostCenter(emp, ec.cost, ec.entries, revenueByCC);
    });
    const costByDepartment = Object.fromEntries(DEPARTMENTS.map((d) => [d.key, 0]));
    employeeCosts.forEach((ec) => { costByDepartment[ec.department] = (costByDepartment[ec.department] || 0) + ec.cost; });
    const byCostCenter = [...COST_CENTERS, UNALLOCATED].map((cc) => {
      const revenue = revenueByCC[cc] || 0;
      const cost = employeeCosts.reduce((s, ec) => s + (ec.byCostCenter?.[cc] || 0), 0);
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
/* Admin: Billing monitor — tracks whether each fee has been invoiced and  */
/* paid. Invoices themselves are issued outside the app; here every fixed  */
/* fee (per client, per month) and every extra fee is a line to follow up. */
/* ---------------------------------------------------------------------- */

const BILLING_OVERDUE_DAYS = 30; // invoiced, unpaid, and older than this → overdue
const BILLING_STATUS_LABELS = { pending: "Not invoiced", invoiced: "Invoiced — unpaid", partial: "Partially paid", overdue: "Overdue", paid: "Paid" };
const BILLING_STATUS_TONE = { pending: "neutral", invoiced: "warn", partial: "warn", overdue: "danger", paid: "accent" };

const monthKeysBetween = (fromKeyStr, toKeyStr) => {
  const out = [];
  let d = fromKey(fromKeyStr.slice(0, 7) + "-01");
  const end = fromKey(toKeyStr.slice(0, 7) + "-01");
  while (d <= end) { out.push(monthKeyOf(d)); d = new Date(d.getFullYear(), d.getMonth() + 1, 1); }
  return out;
};

// Every billable line the firm should have invoiced up to today: one per
// client per month of fixed fee, plus one per extra fee. Derived, never
// stored — the only thing stored is the follow-up status per line.
function billingLines(clients) {
  const todayKey = toKey(TODAY);
  const lines = [];
  clients.forEach((c) => {
    const start = earliestHistoryDate(c.feeHistory);
    if (start) {
      const lastKey = c.endDate && c.endDate < todayKey ? c.endDate : todayKey;
      if (start <= lastKey) {
        monthKeysBetween(start, lastKey).forEach((m) => {
          const fee = feeAsOf(c, m === start.slice(0, 7) ? start : m + "-01").fixedFee;
          if (fee > 0) lines.push({ key: `fixed|${c.id}|${m}`, clientId: c.id, kind: "fixed", month: m, amount: fee, note: "" });
        });
      }
    }
    (c.extraFeeEntries || []).forEach((x) => {
      lines.push({ key: `extra|${x.id}`, clientId: c.id, kind: "extra", month: x.month, amount: x.amount, extraFeeId: x.id, note: x.note || "", costCenter: x.costCenter || "other" });
    });
  });
  return lines;
}

// Every billable line with its follow-up status, paid/outstanding amounts and
// days since invoice — used by the Billing page and the Dashboard alert.
function computeBillingLines(clients, billingStatuses) {
  const statusByKey = Object.fromEntries(billingStatuses.map((s) => [s.kind === "fixed" ? `fixed|${s.clientId}|${s.month}` : `extra|${s.extraFeeId}`, s]));
  const nameOf = (id) => clients.find((c) => c.id === id)?.name || "";
  const vatOf = (id) => { const c = clients.find((x) => x.id === id); return c && c.vatRate !== undefined ? Number(c.vatRate) : 24; };
  return billingLines(clients).map((raw) => {
    // "amount" is the payable (net + VAT); "net" and "vat" are kept for display.
    const rate = vatOf(raw.clientId), vat = Math.round(raw.amount * rate) / 100;
    const l = { ...raw, net: raw.amount, vat, vatRate: rate, amount: raw.amount + vat };
    const st = statusByKey[l.key] || null;
    const status = billingLineStatus(l, st);
    const paidAmount = st?.paidDate ? (st.paidAmount ?? l.amount) : 0;
    const daysSinceInvoice = st?.invoicedDate && !st?.paidDate ? Math.round((TODAY - fromKey(st.invoicedDate)) / 86400000) : null;
    return { ...l, st, status, paidAmount, outstanding: l.amount - paidAmount, daysSinceInvoice };
  }).sort((a, b) => b.month.localeCompare(a.month) || nameOf(a.clientId).localeCompare(nameOf(b.clientId)) || (a.kind === "fixed" ? -1 : 1));
}

function billingLineStatus(line, st) {
  if (!st) return "pending";
  if (st.paidDate) {
    if (st.paidAmount !== null && st.paidAmount < line.amount - 0.005) return "partial";
    return "paid";
  }
  if (st.invoicedDate) {
    const days = Math.round((TODAY - fromKey(st.invoicedDate)) / 86400000);
    return days > BILLING_OVERDUE_DAYS ? "overdue" : "invoiced";
  }
  return "pending";
}

function AdminBilling({ clients, billingStatuses, refetchBilling }) {
  const [viewPeriod, setViewPeriod] = useState("month");
  const [viewAnchor, setViewAnchor] = useState(TODAY);
  const [clientFilter, setClientFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("open");
  const [editing, setEditing] = useState(null); // line being updated
  const [fInvoicedDate, setFInvoicedDate] = useState("");
  const [fInvoiceRef, setFInvoiceRef] = useState("");
  const [fPaidDate, setFPaidDate] = useState("");
  const [fPaidAmount, setFPaidAmount] = useState("");
  const [fNote, setFNote] = useState("");
  const [selected, setSelected] = useState(() => new Set());
  const [allMonths, setAllMonths] = useState(false);
  const [bulkPaidModal, setBulkPaidModal] = useState(false);
  const [bulkPaidDate, setBulkPaidDate] = useState(toKey(TODAY));

  const clientName = (id) => clients.find((c) => c.id === id)?.name || "—";
  const allLines = computeBillingLines(clients, billingStatuses);

  const [rangeStart, rangeEnd] = periodRange(viewPeriod, viewAnchor);
  const fromM = monthKeyOf(rangeStart), toM = monthKeyOf(rangeEnd);
  const inRange = allMonths ? allLines : allLines.filter((l) => l.month >= fromM && l.month <= toM);
  const visible = inRange
    .filter((l) => clientFilter === "all" || l.clientId === clientFilter)
    .filter((l) => statusFilter === "all" || (statusFilter === "open" ? l.status !== "paid" : statusFilter === "unpaid" ? ["invoiced", "partial", "overdue"].includes(l.status) : l.status === statusFilter));

  // Point-in-time totals (all months up to today), independent of the period.
  const sumBy = (arr, pred, f) => arr.filter(pred).reduce((s, l) => s + f(l), 0);
  const toInvoice = sumBy(allLines, (l) => l.status === "pending", (l) => l.amount);
  const invoicedUnpaid = sumBy(allLines, (l) => ["invoiced", "partial", "overdue"].includes(l.status), (l) => l.outstanding);
  const overdue = sumBy(allLines, (l) => l.status === "overdue", (l) => l.outstanding);
  const periodBilled = sumBy(inRange, () => true, (l) => l.amount);
  const periodPaid = sumBy(inRange, () => true, (l) => l.paidAmount);

  const byClient = clients.map((c) => {
    const open = allLines.filter((l) => l.clientId === c.id && l.status !== "paid");
    if (open.length === 0) return null;
    const pending = open.filter((l) => l.status === "pending");
    const unpaid = open.filter((l) => l.status !== "pending");
    const oldestOpen = open.reduce((min, l) => !min || l.month < min ? l.month : min, null);
    const maxDays = unpaid.reduce((m, l) => Math.max(m, l.daysSinceInvoice || 0), 0);
    return { id: c.id, name: c.name, toInvoice: pending.reduce((s, l) => s + l.amount, 0), outstanding: unpaid.reduce((s, l) => s + l.outstanding, 0), openCount: open.length, oldestOpen, maxDays };
  }).filter(Boolean).sort((a, b) => (b.toInvoice + b.outstanding) - (a.toInvoice + a.outstanding));

  function openEdit(line, preset) {
    setEditing(line);
    const st = line.st;
    setFInvoicedDate(st?.invoicedDate || (preset === "invoiced" || preset === "paid" ? toKey(TODAY) : ""));
    setFInvoiceRef(st?.invoiceRef || "");
    setFPaidDate(st?.paidDate || (preset === "paid" ? toKey(TODAY) : ""));
    setFPaidAmount(st?.paidAmount !== null && st?.paidAmount !== undefined ? String(st.paidAmount) : "");
    setFNote(st?.note || "");
  }
  function statusRow(line) {
    return { client_id: line.clientId, kind: line.kind, month: line.month, extra_fee_id: line.kind === "extra" ? line.extraFeeId : null };
  }
  async function saveEdit() {
    if (fPaidDate && !fInvoicedDate) { window.alert("Enter the invoice date as well — a fee can't be paid before it's invoiced."); return; }
    if (fPaidAmount !== "" && Number(fPaidAmount) <= 0) { window.alert("Paid amount must be greater than zero (leave it empty for a full payment)."); return; }
    const row = {
      ...statusRow(editing),
      invoiced_date: fInvoicedDate || null, invoice_ref: fInvoiceRef.trim() || null,
      paid_date: fPaidDate || null, paid_amount: fPaidDate && fPaidAmount !== "" ? Number(fPaidAmount) : null,
      note: fNote.trim() || null, updated_at: new Date().toISOString(),
    };
    const { error } = editing.st
      ? await supabase.from("billing_status").update(row).eq("id", editing.st.id)
      : await supabase.from("billing_status").insert(row);
    if (error) { window.alert(`Couldn't save: ${error.message}`); return; }
    await refetchBilling();
    setEditing(null);
  }
  async function clearStatus() {
    if (!editing.st) { setEditing(null); return; }
    if (!window.confirm("Reset this line to \"Not invoiced\"? The invoice and payment dates will be removed.")) return;
    await supabase.from("billing_status").delete().eq("id", editing.st.id);
    await refetchBilling();
    setEditing(null);
  }
  // Bulk: mark every "Not invoiced" line currently on screen as invoiced today.
  async function markVisibleInvoiced() {
    const pending = visible.filter((l) => l.status === "pending");
    if (pending.length === 0) return;
    if (!window.confirm(`Mark ${pending.length} line${pending.length === 1 ? "" : "s"} (${fmtEur(pending.reduce((s, l) => s + l.amount, 0))}) as invoiced today?`)) return;
    const { error } = await supabase.from("billing_status").insert(pending.map((l) => ({ ...statusRow(l), invoiced_date: toKey(TODAY) })));
    if (error) { window.alert(`Couldn't update: ${error.message}`); return; }
    refetchBilling();
  }

  // ---- Mass actions on the ticked lines ----
  const selectable = visible.filter((l) => l.status !== "paid");
  const selectedLines = allLines.filter((l) => selected.has(l.key) && l.status !== "paid");
  const selectedTotal = selectedLines.reduce((s, l) => s + l.outstanding, 0);
  const selectedPending = selectedLines.filter((l) => l.status === "pending").length;
  const allSelectableTicked = selectable.length > 0 && selectable.every((l) => selected.has(l.key));
  const toggleLine = (key) => setSelected((prev) => { const n = new Set(prev); n.has(key) ? n.delete(key) : n.add(key); return n; });
  const toggleAllVisible = () => setSelected((prev) => { const n = new Set(prev); if (allSelectableTicked) selectable.forEach((l) => n.delete(l.key)); else selectable.forEach((l) => n.add(l.key)); return n; });
  // Quick select looks at every month (so old unpaid lines are included) and
  // switches the table to "All months" so the ticked lines are visible.
  const selectByStatus = (statuses) => {
    setAllMonths(true);
    setSelected(new Set(allLines.filter((l) => (clientFilter === "all" || l.clientId === clientFilter) && statuses.includes(l.status)).map((l) => l.key)));
  };

  // Marks every ticked line as paid in full on one date. Lines that were
  // never invoiced are invoiced on the same date, so the record stays consistent.
  async function bulkMarkPaid() {
    if (!bulkPaidDate) { window.alert("Choose the payment date."); return; }
    const now = new Date().toISOString();
    const withInvoice = selectedLines.filter((l) => l.st?.invoicedDate);
    const withRowNoInvoice = selectedLines.filter((l) => l.st && !l.st.invoicedDate);
    const noRow = selectedLines.filter((l) => !l.st);
    const errors = [];
    if (withInvoice.length) {
      const { error } = await supabase.from("billing_status").update({ paid_date: bulkPaidDate, paid_amount: null, updated_at: now }).in("id", withInvoice.map((l) => l.st.id));
      if (error) errors.push(error.message);
    }
    if (withRowNoInvoice.length) {
      const { error } = await supabase.from("billing_status").update({ invoiced_date: bulkPaidDate, paid_date: bulkPaidDate, paid_amount: null, updated_at: now }).in("id", withRowNoInvoice.map((l) => l.st.id));
      if (error) errors.push(error.message);
    }
    if (noRow.length) {
      const { error } = await supabase.from("billing_status").insert(noRow.map((l) => ({ ...statusRow(l), invoiced_date: bulkPaidDate, paid_date: bulkPaidDate })));
      if (error) errors.push(error.message);
    }
    if (errors.length) window.alert(`Some lines could not be updated: ${errors.join("; ")}`);
    await refetchBilling();
    setSelected(new Set());
    setBulkPaidModal(false);
  }
  // Marks the ticked lines that are still "Not invoiced" as invoiced today.
  async function bulkMarkInvoiced() {
    const pending = selectedLines.filter((l) => l.status === "pending");
    if (pending.length === 0) return;
    const { error } = await supabase.from("billing_status").insert(pending.map((l) => ({ ...statusRow(l), invoiced_date: toKey(TODAY) })));
    if (error) { window.alert(`Couldn't update: ${error.message}`); return; }
    await refetchBilling();
    setSelected(new Set());
  }

  // One row per line still to be invoiced, coded the way myDATA expects it,
  // ready to be keyed into (or imported by) the invoicing program.
  function invoicingRows() {
    const monthGr = (m) => { const [y, mo] = m.split("-").map(Number); return new Date(y, mo - 1, 1).toLocaleDateString("el-GR", { month: "long", year: "numeric" }); };
    return visible.filter((l) => l.status === "pending").map((l) => {
      const c = clients.find((x) => x.id === l.clientId) || {};
      const type = MYDATA_INVOICE_TYPES[c.invoiceType] ? c.invoiceType : "2.1";
      const t = MYDATA_INVOICE_TYPES[type];
      const svcKey = l.kind === "fixed" ? (c.fixedService || "accounting") : (String(l.note || "").startsWith("Start-up & business formation") ? "formation" : EXTRA_SERVICE_BY_CC[l.costCenter] || "other");
      const svc = INVOICE_SERVICES[svcKey] || INVOICE_SERVICES.other;
      const noVat = Number(l.vatRate) === 0;
      const ex = noVat ? MYDATA_VAT_EXEMPTIONS.find(([k]) => k === c.vatExemptionCategory) : null;
      const issues = [!c.afm && "Λείπει ΑΦΜ", noVat && !ex && "Λείπει αιτία εξαίρεσης ΦΠΑ"].filter(Boolean).join("; ");
      return {
        "Πελάτης": c.name || "", "ΑΦΜ": c.afm || "", "Email": c.email || "",
        "Τύπος παραστατικού (κωδ.)": type, "Τύπος παραστατικού": t.label,
        "Περιγραφή υπηρεσίας": l.kind === "fixed" ? `${svc} — ${monthGr(l.month)}` : `${svc}${l.note ? ` — ${l.note}` : ""} (${monthGr(l.month)})`,
        "Περίοδος": l.month, "Καθαρή αξία (€)": Number(l.net.toFixed(2)),
        "Κατηγορία ΦΠΑ (κωδ.)": MYDATA_VAT_CATEGORY[noVat ? 0 : 24] ?? "", "ΦΠΑ %": l.vatRate, "ΦΠΑ (€)": Number(l.vat.toFixed(2)), "Σύνολο (€)": Number(l.amount.toFixed(2)),
        "Αιτία εξαίρεσης ΦΠΑ (κωδ.)": ex ? ex[0] : "", "Αιτία εξαίρεσης ΦΠΑ": ex ? ex[1] : "",
        "Χαρακτηρισμός εσόδων — κατηγορία": MYDATA_INCOME_CATEGORY.code, "Χαρακτηρισμός εσόδων — τύπος": t.incomeType, "Περιγραφή τύπου": t.incomeTypeLabel,
        "Έλεγχος": issues,
      };
    });
  }

  function exportBilling() {
    const toInvoiceRows = invoicingRows();
    exportToExcel(`billing-monitor-${toKey(TODAY)}.xlsx`, [
      ...(toInvoiceRows.length ? [{ name: "To invoice (myDATA)", rows: toInvoiceRows }] : []),
      { name: "Lines", rows: visible.map((l) => ({
        Month: l.month, Client: clientName(l.clientId), Type: l.kind === "fixed" ? "Fixed fee" : "Extra fee", Note: l.note,
        "Net (€)": l.net, "VAT %": l.vatRate, "VAT (€)": l.vat, "Payable (€)": Number(l.amount.toFixed(2)), "Invoiced on": l.st?.invoicedDate || "", "Invoice ref": l.st?.invoiceRef || "",
        "Paid on": l.st?.paidDate || "", "Paid (€)": Number(l.paidAmount.toFixed(2)), "Outstanding (€)": Number(l.outstanding.toFixed(2)),
        Status: BILLING_STATUS_LABELS[l.status], "Days since invoice": l.daysSinceInvoice ?? "",
      })) },
      { name: "By client", rows: byClient.map((c) => ({ Client: c.name, "To invoice (€)": Number(c.toInvoice.toFixed(2)), "Invoiced, unpaid (€)": Number(c.outstanding.toFixed(2)), "Open lines": c.openCount, "Oldest open month": c.oldestOpen, "Max days since invoice": c.maxDays })) },
    ]);
  }

  const pendingVisibleCount = visible.filter((l) => l.status === "pending").length;

  return (
    <div>
      <PageHeader title="Billing" sub="Follow-up of invoicing and settlement for every fixed and extra fee. Invoices are issued outside the app — here you only record when each fee was invoiced and when it was paid."
        right={
          <div style={{ display: "flex", gap: 8 }}>
            <Btn variant="secondary" onClick={exportBilling}>Export to Excel</Btn>
            <Btn icon={Check} onClick={markVisibleInvoiced} disabled={pendingVisibleCount === 0}>Mark {pendingVisibleCount || ""} pending as invoiced</Btn>
          </div>
        } />

      <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
        <StatCard label="Not invoiced yet" value={fmtEur(toInvoice)} accent={toInvoice > 0 ? C.warn : C.accent} sub="All months to date · incl. VAT" />
        <StatCard label="Invoiced, unpaid" value={fmtEur(invoicedUnpaid)} accent={invoicedUnpaid > 0 ? C.warn : C.accent} sub="Outstanding as of today · incl. VAT" />
        <StatCard label="Overdue" value={fmtEur(overdue)} accent={overdue > 0 ? C.danger : C.accent} sub={`Invoiced > ${BILLING_OVERDUE_DAYS} days ago, unpaid`} />
        <StatCard label="Fees in period" value={fmtEur(periodBilled)} sub={`${periodLabel(viewPeriod, rangeStart, rangeEnd)} · incl. VAT`} />
        <StatCard label="Paid, in period" value={fmtEur(periodPaid)} accent={C.accent} sub="Of the fees in this period" />
      </div>

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 12 }}>Open items by client</div>
        {byClient.length === 0 ? (
          <EmptyState icon={Wallet} title="Nothing open" sub="Every fee to date is invoiced and paid." />
        ) : (
          <TableShell headers={["Client", "Not invoiced", "Invoiced, unpaid", "Open lines", "Oldest open month", "Longest wait", ""]}>
            {byClient.map((c) => (
              <tr key={c.id}>
                <Td>{c.name}</Td>
                <Td mono style={{ color: c.toInvoice > 0 ? C.warn : C.inkMuted }}>{c.toInvoice > 0 ? fmtEur(c.toInvoice) : "—"}</Td>
                <Td mono style={{ fontWeight: 700, color: c.maxDays > BILLING_OVERDUE_DAYS ? C.danger : c.outstanding > 0 ? C.warn : C.inkMuted }}>{c.outstanding > 0 ? fmtEur(c.outstanding) : "—"}</Td>
                <Td mono>{c.openCount}</Td>
                <Td mono style={{ color: C.inkMuted }}>{c.oldestOpen}</Td>
                <Td mono style={{ color: c.maxDays > BILLING_OVERDUE_DAYS ? C.danger : C.inkMuted }}>{c.maxDays > 0 ? `${c.maxDays}d` : "—"}</Td>
                <Td right><Btn size="sm" variant="ghost" onClick={() => { setClientFilter(c.id); setStatusFilter("open"); setViewPeriod("year"); }}>Show</Btn></Td>
              </tr>
            ))}
          </TableShell>
        )}
      </Panel>

      <Panel>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: 12, flexWrap: "wrap", gap: 10 }}>
          <div>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 8 }}>Fee lines</div>
            <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", opacity: allMonths ? 0.5 : 1 }}>
              <PeriodSelector period={viewPeriod} setPeriod={setViewPeriod} anchor={viewAnchor} setAnchor={setViewAnchor} />
            </div>
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontFamily: sans, fontSize: 12.5, color: C.ink, marginTop: 8, cursor: "pointer" }}>
              <input type="checkbox" checked={allMonths} onChange={(e) => setAllMonths(e.target.checked)} /> Show all months (ignore the period above)
            </label>
          </div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <Field label="Client">
              <select style={{ ...inputStyle, width: 190 }} value={clientFilter} onChange={(e) => setClientFilter(e.target.value)}>
                <option value="all">All clients</option>
                {clients.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
            </Field>
            <Field label="Status">
              <select style={{ ...inputStyle, width: 190 }} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
                <option value="open">Open (not paid)</option>
                <option value="unpaid">Invoiced, unpaid + overdue</option>
                <option value="all">All</option>
                {Object.keys(BILLING_STATUS_LABELS).map((k) => <option key={k} value={k}>{BILLING_STATUS_LABELS[k]}</option>)}
              </select>
            </Field>
          </div>
        </div>
        {visible.length === 0 ? (
          <EmptyState icon={Receipt} title="No fee lines match" sub="Change the period or filters." />
        ) : (
          <>
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 10 }}>
            <span style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted }}>Quick select:</span>
            <Btn size="sm" variant="ghost" onClick={() => selectByStatus(["overdue"])}>Overdue</Btn>
            <Btn size="sm" variant="ghost" onClick={() => selectByStatus(["invoiced", "partial", "overdue"])}>All invoiced, unpaid</Btn>
            <Btn size="sm" variant="ghost" onClick={() => selectByStatus(["pending"])}>Not invoiced</Btn>
            {selected.size > 0 && <Btn size="sm" variant="ghost" onClick={() => setSelected(new Set())}>Clear</Btn>}
            {selectedLines.length > 0 && (
              <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center", background: C.accentSoft, borderRadius: 8, padding: "6px 10px" }}>
                <span style={{ fontFamily: sans, fontSize: 12.5, color: C.ink }}><strong>{selectedLines.length}</strong> selected · {fmtEur(selectedTotal)}</span>
                {selectedPending > 0 && <Btn size="sm" variant="secondary" icon={Receipt} onClick={bulkMarkInvoiced}>Mark {selectedPending} invoiced</Btn>}
                <Btn size="sm" icon={Wallet} onClick={() => { setBulkPaidDate(toKey(TODAY)); setBulkPaidModal(true); }}>Mark selected as paid</Btn>
              </div>
            )}
          </div>
          <TableShell headers={[<input key="all" type="checkbox" checked={allSelectableTicked} onChange={toggleAllVisible} disabled={selectable.length === 0} title="Select all open lines on screen" />, "Month", "Client", "Type", "Net", "VAT", "Payable", "Invoiced", "Paid", "Status", ""]}>
            {visible.map((l) => (
              <tr key={l.key} style={{ background: selected.has(l.key) ? C.accentSoft : undefined }}>
                <Td>{l.status !== "paid" ? <input type="checkbox" checked={selected.has(l.key)} onChange={() => toggleLine(l.key)} /> : null}</Td>
                <Td mono>{l.month}</Td>
                <Td>{clientName(l.clientId)}</Td>
                <Td style={{ color: C.inkMuted }} title={l.note || undefined}>{l.kind === "fixed" ? "Fixed fee" : `Extra${l.note ? ` — ${l.note}` : ""}`}</Td>
                <Td mono style={{ color: C.inkMuted }}>{fmtEur(l.net)}</Td>
                <Td mono style={{ color: C.inkMuted }} title={`ΦΠΑ ${l.vatRate}%`}>{l.vat > 0 ? fmtEur(l.vat) : "0%"}</Td>
                <Td mono style={{ fontWeight: 700 }}>{fmtEur(l.amount)}</Td>
                <Td mono style={{ color: l.st?.invoicedDate ? C.ink : C.inkFaint }}>
                  {l.st?.invoicedDate ? <>{l.st.invoicedDate}{l.st.invoiceRef ? <span style={{ color: C.inkMuted }}> · {l.st.invoiceRef}</span> : null}</> : "—"}
                </Td>
                <Td mono style={{ color: l.st?.paidDate ? C.ink : C.inkFaint }}>
                  {l.st?.paidDate ? <>{l.st.paidDate}{l.status === "partial" ? <span style={{ color: C.warn }}> · {fmtEur(l.paidAmount)}</span> : null}</> : "—"}
                </Td>
                <Td>
                  <Badge tone={BILLING_STATUS_TONE[l.status]}>{BILLING_STATUS_LABELS[l.status]}</Badge>
                  {l.daysSinceInvoice !== null && l.status !== "paid" ? <span style={{ fontFamily: mono, fontSize: 11, color: l.status === "overdue" ? C.danger : C.inkMuted, marginLeft: 6 }}>{l.daysSinceInvoice}d</span> : null}
                </Td>
                <Td right>
                  <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                    {l.status === "pending" && <Btn size="sm" variant="secondary" icon={Receipt} onClick={() => openEdit(l, "invoiced")}>Invoiced</Btn>}
                    {l.status !== "paid" && l.status !== "pending" && <Btn size="sm" variant="secondary" icon={Wallet} onClick={() => openEdit(l, "paid")}>Paid</Btn>}
                    <Btn size="sm" variant="ghost" icon={Pencil} onClick={() => openEdit(l)}>Edit</Btn>
                  </div>
                </Td>
              </tr>
            ))}
            <tr style={{ fontWeight: 700, borderTop: `2px solid ${C.borderStrong}` }}>
              <Td /><Td>Total</Td><Td /><Td />
              <Td mono>{fmtEur(visible.reduce((s, l) => s + l.net, 0))}</Td>
              <Td mono>{fmtEur(visible.reduce((s, l) => s + l.vat, 0))}</Td>
              <Td mono>{fmtEur(visible.reduce((s, l) => s + l.amount, 0))}</Td>
              <Td mono>{fmtEur(visible.filter((l) => l.st?.invoicedDate).reduce((s, l) => s + l.amount, 0))}</Td>
              <Td mono>{fmtEur(visible.reduce((s, l) => s + l.paidAmount, 0))}</Td>
              <Td /><Td />
            </tr>
          </TableShell>
          </>
        )}
      </Panel>

      {bulkPaidModal && (
        <Modal title={`Mark ${selectedLines.length} line${selectedLines.length === 1 ? "" : "s"} as paid`} onClose={() => setBulkPaidModal(false)} width={440}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted, background: C.accentSoft, borderRadius: 7, padding: "8px 10px" }}>
              Total outstanding on the selected lines: <strong>{fmtEur(selectedTotal)}</strong>. Each line is recorded as paid in full on the date below.
              {selectedPending > 0 ? ` ${selectedPending} line(s) were never marked as invoiced — they are recorded as invoiced on the same date.` : ""}
            </div>
            <Field label="Payment date"><input type="date" style={inputStyle} value={bulkPaidDate} onChange={(e) => setBulkPaidDate(e.target.value)} autoFocus /></Field>
            <div style={{ maxHeight: 170, overflowY: "auto", border: `1px solid ${C.border}`, borderRadius: 7, padding: "4px 8px" }}>
              {selectedLines.map((l) => (
                <div key={l.key} style={{ display: "flex", justifyContent: "space-between", fontFamily: sans, fontSize: 12, padding: "2px 0", color: C.ink }}>
                  <span>{l.month} · {clientName(l.clientId)} · {l.kind === "fixed" ? "Fixed" : "Extra"}</span><span style={{ fontFamily: mono }}>{fmtEur(l.outstanding)}</span>
                </div>
              ))}
            </div>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <Btn variant="secondary" onClick={() => setBulkPaidModal(false)}>Cancel</Btn>
              <Btn icon={Check} onClick={bulkMarkPaid}>Mark as paid</Btn>
            </div>
          </div>
        </Modal>
      )}

      {editing && (
        <Modal title={`${clientName(editing.clientId)} — ${editing.kind === "fixed" ? "Fixed fee" : "Extra fee"} ${editing.month}`} onClose={() => setEditing(null)} width={440}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted, background: C.accentSoft, borderRadius: 7, padding: "8px 10px" }}>
              Amount <strong>{fmtEur(editing.amount)}</strong>{editing.note ? ` · ${editing.note}` : ""}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <Field label="Invoiced on">
                <input type="date" style={inputStyle} value={fInvoicedDate} onChange={(e) => setFInvoicedDate(e.target.value)} />
              </Field>
              <Field label="Invoice ref. (optional)">
                <input style={inputStyle} value={fInvoiceRef} onChange={(e) => setFInvoiceRef(e.target.value)} placeholder="e.g. ΤΠΥ 142" />
              </Field>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <Field label="Paid on">
                <input type="date" style={inputStyle} value={fPaidDate} onChange={(e) => setFPaidDate(e.target.value)} />
              </Field>
              <Field label="Paid amount (empty = full)">
                <input type="number" min="0" style={inputStyle} value={fPaidAmount} onChange={(e) => setFPaidAmount(e.target.value)} placeholder={String(editing.amount)} />
              </Field>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              {!fInvoicedDate && <Btn variant="ghost" size="sm" onClick={() => setFInvoicedDate(toKey(TODAY))}>Invoiced today</Btn>}
              {!fPaidDate && <Btn variant="ghost" size="sm" onClick={() => { if (!fInvoicedDate) setFInvoicedDate(toKey(TODAY)); setFPaidDate(toKey(TODAY)); }}>Paid today</Btn>}
            </div>
            <Field label="Note (optional)">
              <input style={inputStyle} value={fNote} onChange={(e) => setFNote(e.target.value)} placeholder="e.g. Client asked for split payment" />
            </Field>
            <div style={{ display: "flex", gap: 8, justifyContent: "space-between", marginTop: 6 }}>
              <Btn variant="ghost" size="sm" onClick={clearStatus} disabled={!editing.st}>Reset to not invoiced</Btn>
              <div style={{ display: "flex", gap: 8 }}>
                <Btn variant="secondary" onClick={() => setEditing(null)}>Cancel</Btn>
                <Btn onClick={saveEdit}>Save</Btn>
              </div>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* CRM: shared task & note editors                                         */
/* ---------------------------------------------------------------------- */

const textareaStyle = { ...inputStyle, minHeight: 80, resize: "vertical", lineHeight: 1.45 };

// Create/edit a task. `task` null = new. `fixedClientId` pins the client
// (from a client card); otherwise the client is picked in the form.
function TaskModal({ user, task, fixedClientId, clients, employees, onClose, onSaved }) {
  const [title, setTitle] = useState(task?.title || "");
  const [description, setDescription] = useState(task?.description || "");
  const [clientId, setClientId] = useState(task?.clientId || fixedClientId || "");
  const [assignedTo, setAssignedTo] = useState(task?.assignedTo || user.id);
  const [dueDate, setDueDate] = useState(task?.dueDate || "");
  const [priority, setPriority] = useState(task?.priority || "normal");
  const [status, setStatus] = useState(task?.status || "open");

  async function save() {
    if (!title.trim()) { window.alert("Give the task a title."); return; }
    const row = {
      title: title.trim(), description: description.trim() || null, client_id: clientId || null,
      assigned_to: assignedTo || null, due_date: dueDate || null, priority, status,
      done_at: status === "done" ? (task?.doneAt || toKey(TODAY)) : null,
    };
    const { error } = task
      ? await supabase.from("tasks").update(row).eq("id", task.id)
      : await supabase.from("tasks").insert({ ...row, created_by: user.id });
    if (error) { window.alert(`Couldn't save task: ${error.message}`); return; }
    await onSaved();
    onClose();
  }
  async function remove() {
    if (!window.confirm("Delete this task?")) return;
    const { error } = await supabase.from("tasks").delete().eq("id", task.id);
    if (error) { window.alert(`Couldn't delete: ${error.message}`); return; }
    await onSaved();
    onClose();
  }
  const activeEmployees = employees.filter((e) => e.active || e.id === assignedTo);
  return (
    <Modal title={task ? "Edit task" : "New task"} onClose={onClose} width={460}>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="Title">
          <input style={inputStyle} value={title} onChange={(e) => setTitle(e.target.value)} placeholder="e.g. Send ΦΠΑ Q3 draft for approval" autoFocus />
        </Field>
        <Field label="Details (optional)">
          <textarea style={textareaStyle} value={description} onChange={(e) => setDescription(e.target.value)} />
        </Field>
        {!fixedClientId && (
          <Field label="Client (optional)">
            <select style={inputStyle} value={clientId} onChange={(e) => setClientId(e.target.value)}>
              <option value="">— no client —</option>
              {clients.filter((c) => c.active || c.id === clientId).map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
        )}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <Field label="Assigned to">
            <select style={inputStyle} value={assignedTo || ""} onChange={(e) => setAssignedTo(e.target.value)}>
              <option value="">— unassigned —</option>
              {activeEmployees.map((e) => <option key={e.id} value={e.id}>{e.name}{e.id === user.id ? " (me)" : ""}</option>)}
            </select>
          </Field>
          <Field label="Due date">
            <input type="date" style={inputStyle} value={dueDate} onChange={(e) => setDueDate(e.target.value)} />
          </Field>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <Field label="Priority">
            <select style={inputStyle} value={priority} onChange={(e) => setPriority(e.target.value)}>
              {Object.entries(TASK_PRIORITIES).map(([k, l]) => <option key={k} value={k}>{l}</option>)}
            </select>
          </Field>
          <Field label="Status">
            <select style={inputStyle} value={status} onChange={(e) => setStatus(e.target.value)}>
              {Object.entries(TASK_STATUSES).map(([k, l]) => <option key={k} value={k}>{l}</option>)}
            </select>
          </Field>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6 }}>
          <div>{task && (task.createdBy === user.id || user.role === "admin") && <Btn variant="ghost" size="sm" icon={Trash2} onClick={remove}>Delete</Btn>}</div>
          <div style={{ display: "flex", gap: 8 }}>
            <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
            <Btn onClick={save}>{task ? "Save" : "Create task"}</Btn>
          </div>
        </div>
      </div>
    </Modal>
  );
}

// Add/edit a communication note (call, email, meeting, note) on a client.
function NoteModal({ user, note, fixedClientId, clients, onClose, onSaved }) {
  const [clientId, setClientId] = useState(note?.clientId || fixedClientId || "");
  const [kind, setKind] = useState(note?.kind || "note");
  const [noteDate, setNoteDate] = useState(note?.noteDate || toKey(TODAY));
  const [content, setContent] = useState(note?.content || "");
  async function save() {
    if (!clientId) { window.alert("Choose a client."); return; }
    if (!content.trim()) { window.alert("Write the note."); return; }
    const row = { client_id: clientId, kind, note_date: noteDate || toKey(TODAY), content: content.trim() };
    const { error } = note
      ? await supabase.from("client_notes").update(row).eq("id", note.id)
      : await supabase.from("client_notes").insert({ ...row, author_id: user.id });
    if (error) { window.alert(`Couldn't save note: ${error.message}`); return; }
    await onSaved();
    onClose();
  }
  async function remove() {
    if (!window.confirm("Delete this note?")) return;
    await supabase.from("client_notes").delete().eq("id", note.id);
    await onSaved();
    onClose();
  }
  return (
    <Modal title={note ? "Edit note" : "Add note"} onClose={onClose} width={460}>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        {!fixedClientId && (
          <Field label="Client">
            <select style={inputStyle} value={clientId} onChange={(e) => setClientId(e.target.value)}>
              <option value="">— choose —</option>
              {clients.filter((c) => c.active || c.id === clientId).map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
        )}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <Field label="Type">
            <select style={inputStyle} value={kind} onChange={(e) => setKind(e.target.value)}>
              {Object.entries(NOTE_KINDS).map(([k, l]) => <option key={k} value={k}>{l}</option>)}
            </select>
          </Field>
          <Field label="Date">
            <input type="date" style={inputStyle} value={noteDate} onChange={(e) => setNoteDate(e.target.value)} />
          </Field>
        </div>
        <Field label="Note">
          <textarea style={{ ...textareaStyle, minHeight: 110 }} value={content} onChange={(e) => setContent(e.target.value)} placeholder="What was discussed / agreed / sent" autoFocus />
        </Field>
        <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6 }}>
          <div>{note && (note.authorId === user.id || user.role === "admin") && <Btn variant="ghost" size="sm" icon={Trash2} onClick={remove}>Delete</Btn>}</div>
          <div style={{ display: "flex", gap: 8 }}>
            <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
            <Btn onClick={save}>{note ? "Save" : "Add note"}</Btn>
          </div>
        </div>
      </div>
    </Modal>
  );
}

// One task list used in three places: the client card, the admin Tasks
// page and the employee's My Tasks page.
function TaskList({ user, tasks, clients, employees, onEdit, onQuickStatus, showClient = true, showAssignee = true, onOpenCard }) {
  const empName = (id) => employees.find((e) => e.id === id)?.name || "Unassigned";
  const clientLabel = (id) => clients.find((c) => c.id === id)?.name || "—";
  const sorted = [...tasks].sort((a, b) => {
    const rank = (t) => t.status === "done" ? 3 : taskIsOverdue(t) ? 0 : t.dueDate ? 1 : 2;
    return rank(a) - rank(b) || (a.dueDate || "9999").localeCompare(b.dueDate || "9999") || (a.priority === "high" ? -1 : 1);
  });
  if (sorted.length === 0) return <EmptyState icon={ListChecks} title="No tasks" />;
  return (
    <TableShell headers={["Due", "Task", ...(showClient ? ["Client"] : []), ...(showAssignee ? ["Assigned to"] : []), "Priority", "Status", ""]}>
      {sorted.map((t) => {
        const overdue = taskIsOverdue(t);
        return (
          <tr key={t.id} style={{ opacity: t.status === "done" ? 0.55 : 1 }}>
            <Td mono style={{ color: overdue ? C.danger : t.dueDate ? undefined : C.inkFaint }}>{t.dueDate || "—"}</Td>
            <Td title={t.description || undefined}>
              <div style={{ fontWeight: 600 }}>{t.kind === "onboarding" && <span style={{ fontFamily: mono, fontSize: 10, color: C.warn, marginRight: 6 }}>{(t.stage || "setup").toUpperCase()}</span>}{t.title}</div>
              {t.description && <div style={{ fontSize: 11.5, color: C.inkMuted, marginTop: 2, maxWidth: 420, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{t.description}</div>}
            </Td>
            {showClient && <Td>{t.clientId ? (onOpenCard ? <button onClick={() => onOpenCard(t.clientId)} style={{ ...iconBtnStyle, padding: 0, fontFamily: sans, fontSize: 13, color: C.accentDark, textDecoration: "underline" }}>{clientLabel(t.clientId)}</button> : clientLabel(t.clientId)) : <span style={{ color: C.inkFaint }}>—</span>}</Td>}
            {showAssignee && <Td>{empName(t.assignedTo)}{t.assignedTo === user.id ? <span style={{ color: C.inkFaint }}> (me)</span> : null}</Td>}
            <Td><Badge tone={TASK_PRIORITY_TONE[t.priority]}>{TASK_PRIORITIES[t.priority]}</Badge></Td>
            <Td><Badge tone={overdue ? "danger" : TASK_STATUS_TONE[t.status]}>{overdue ? "Overdue" : TASK_STATUSES[t.status]}</Badge></Td>
            <Td right>
              <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                {t.status === "open" && <Btn size="sm" variant="ghost" onClick={() => onQuickStatus(t, "in_progress")}>Start</Btn>}
                {t.status !== "done" && <Btn size="sm" variant="secondary" icon={Check} onClick={() => onQuickStatus(t, "done")}>Done</Btn>}
                {t.status === "done" && <Btn size="sm" variant="ghost" onClick={() => onQuickStatus(t, "open")}>Reopen</Btn>}
                <Btn size="sm" variant="ghost" icon={Pencil} onClick={() => onEdit(t)}>Edit</Btn>
              </div>
            </Td>
          </tr>
        );
      })}
    </TableShell>
  );
}

/* ---------------------------------------------------------------------- */
/* Tasks page — admin ("all") sees every task; employees ("mine") see the  */
/* tasks assigned to them and can create tasks and client notes.          */
/* ---------------------------------------------------------------------- */

function TasksPage({ user, mode, tasks, clients, employees, notes, refetchTasks, refetchNotes, onOpenCard, templates = [], refetchTemplates }) {
  const [showTemplates, setShowTemplates] = useState(false);
  const [tplEditing, setTplEditing] = useState(null); // null | "new" | template
  const [tplTitle, setTplTitle] = useState(""); const [tplDesc, setTplDesc] = useState(""); const [tplDays, setTplDays] = useState("7"); const [tplOrder, setTplOrder] = useState(""); const [tplStage, setTplStage] = useState("setup"); const [tplApplies, setTplApplies] = useState("all");
  function openTpl(t) { setTplEditing(t || "new"); setTplTitle(t?.title || ""); setTplDesc(t?.description || ""); setTplDays(String(t?.dueDays ?? 7)); setTplOrder(String(t?.sortOrder ?? (templates.length + 1))); setTplStage(t?.stage || "setup"); setTplApplies(t?.appliesTo || "all"); }
  async function saveTpl() {
    if (!tplTitle.trim()) { window.alert("Enter a title."); return; }
    const row = { title: tplTitle.trim(), description: tplDesc.trim() || null, due_days: Number(tplDays) || 0, sort_order: Number(tplOrder) || 0, stage: tplStage, applies_to: tplApplies };
    const { error } = tplEditing === "new" ? await supabase.from("onboarding_templates").insert(row) : await supabase.from("onboarding_templates").update(row).eq("id", tplEditing.id);
    if (error) { window.alert(`Couldn't save: ${error.message}`); return; }
    await refetchTemplates(); setTplEditing(null);
  }
  async function toggleTpl(t) { await supabase.from("onboarding_templates").update({ active: !t.active }).eq("id", t.id); refetchTemplates(); }
  async function deleteTpl(t) { if (!window.confirm(`Delete "${t.title}" from the checklist? Existing clients' tasks are not affected.`)) return; await supabase.from("onboarding_templates").delete().eq("id", t.id); refetchTemplates(); }
  const [assigneeFilter, setAssigneeFilter] = useState(mode === "mine" ? user.id : "all");
  const [clientFilter, setClientFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("open");
  const [taskModal, setTaskModal] = useState(null); // null | "new" | task
  const [noteModal, setNoteModal] = useState(false);

  const visible = tasks
    .filter((t) => mode === "all" || t.assignedTo === user.id || t.createdBy === user.id)
    .filter((t) => assigneeFilter === "all" || (assigneeFilter === "unassigned" ? !t.assignedTo : t.assignedTo === assigneeFilter))
    .filter((t) => clientFilter === "all" || t.clientId === clientFilter)
    .filter((t) => statusFilter === "all" || (statusFilter === "open" ? t.status !== "done" : statusFilter === "overdue" ? taskIsOverdue(t) : t.status === statusFilter));

  async function quickStatus(t, status) {
    const { error } = await supabase.from("tasks").update({ status, done_at: status === "done" ? toKey(TODAY) : null }).eq("id", t.id);
    if (error) { window.alert(`Couldn't update: ${error.message}`); return; }
    refetchTasks();
  }
  const myOpen = tasks.filter((t) => t.assignedTo === user.id && t.status !== "done");
  const myOverdue = myOpen.filter(taskIsOverdue);
  const recentNotes = mode === "mine" ? notes.filter((n) => n.authorId === user.id).sort((a, b) => b.noteDate.localeCompare(a.noteDate)).slice(0, 8) : [];

  return (
    <div>
      <PageHeader title={mode === "all" ? "Tasks" : "My Tasks"} sub={mode === "all" ? "Every task across clients and people." : "Tasks assigned to you, plus anything you created. You can also log a client note from here."}
        right={
          <div style={{ display: "flex", gap: 8 }}>
            {mode === "all" && <Btn variant="secondary" icon={ListChecks} onClick={() => setShowTemplates((v) => !v)}>{showTemplates ? "Hide checklist" : "Onboarding checklist"}</Btn>}
            <Btn variant="secondary" icon={MessageSquare} onClick={() => setNoteModal(true)}>Add client note</Btn>
            <Btn icon={Plus} onClick={() => setTaskModal("new")}>New task</Btn>
          </div>
        } />

      {mode === "all" && showTemplates && (
        <Panel style={{ marginBottom: 18 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <div>
              <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink }}>Onboarding checklist template</div>
              <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted }}>Created for every client when onboarding starts (the "Company formation" stage only when you say the company is newly formed). Due date = start + days. The client goes live when all of them are done.</div>
            </div>
            <Btn size="sm" icon={Plus} onClick={() => openTpl(null)}>Add item</Btn>
          </div>
          <TableShell headers={["Stage", "#", "Task", "Details", "Applies to", "Due (days)", "Active", ""]}>
            {[...templates].sort((a, b) => ONBOARDING_STAGES.findIndex((x) => x.key === a.stage) - ONBOARDING_STAGES.findIndex((x) => x.key === b.stage) || a.sortOrder - b.sortOrder).map((t) => (
              <tr key={t.id} style={{ opacity: t.active ? 1 : 0.5 }}>
                <Td style={{ color: C.inkMuted, fontSize: 12 }}>{stageLabel(t.stage)}</Td>
                <Td mono>{t.sortOrder}</Td>
                <Td style={{ fontWeight: 600 }}>{t.title}</Td>
                <Td style={{ color: C.inkMuted }}>{t.description || "—"}</Td>
                <Td style={{ fontSize: 12, color: t.appliesTo === "all" ? C.inkMuted : C.warn }}>{APPLIES_TO_LABELS[t.appliesTo] || t.appliesTo}</Td>
                <Td mono>+{t.dueDays}</Td>
                <Td>{t.active ? <Badge tone="accent">Active</Badge> : <Badge tone="neutral">Off</Badge>}</Td>
                <Td right>
                  <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                    <Btn size="sm" variant="ghost" icon={Pencil} onClick={() => openTpl(t)}>Edit</Btn>
                    <Btn size="sm" variant="ghost" onClick={() => toggleTpl(t)}>{t.active ? "Disable" : "Enable"}</Btn>
                    <Btn size="sm" variant="ghost" icon={Trash2} onClick={() => deleteTpl(t)}>Delete</Btn>
                  </div>
                </Td>
              </tr>
            ))}
          </TableShell>
          {tplEditing && (
            <Modal title={tplEditing === "new" ? "Add checklist item" : "Edit checklist item"} onClose={() => setTplEditing(null)} width={440}>
              <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                <Field label="Stage">
                  <select style={inputStyle} value={tplStage} onChange={(e) => setTplStage(e.target.value)}>
                    {ONBOARDING_STAGES.map((st) => <option key={st.key} value={st.key}>{st.label}</option>)}
                  </select>
                </Field>
                <Field label="Applies to">
                  <select style={inputStyle} value={tplApplies} onChange={(e) => setTplApplies(e.target.value)}>
                    {Object.entries(APPLIES_TO_LABELS).map(([k, l]) => <option key={k} value={k}>{l}</option>)}
                  </select>
                </Field>
                <Field label="Task title"><input style={inputStyle} value={tplTitle} onChange={(e) => setTplTitle(e.target.value)} autoFocus /></Field>
                <Field label="Details (optional)"><input style={inputStyle} value={tplDesc} onChange={(e) => setTplDesc(e.target.value)} /></Field>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                  <Field label="Due: days after start"><input type="number" min="0" style={inputStyle} value={tplDays} onChange={(e) => setTplDays(e.target.value)} /></Field>
                  <Field label="Order"><input type="number" min="0" style={inputStyle} value={tplOrder} onChange={(e) => setTplOrder(e.target.value)} /></Field>
                </div>
                <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                  <Btn variant="secondary" onClick={() => setTplEditing(null)}>Cancel</Btn>
                  <Btn onClick={saveTpl}>Save</Btn>
                </div>
              </div>
            </Modal>
          )}
        </Panel>
      )}

      {mode === "mine" && (
        <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
          <StatCard label="My open tasks" value={String(myOpen.length)} sub="Assigned to me, not done" />
          <StatCard label="Overdue" value={String(myOverdue.length)} accent={myOverdue.length > 0 ? C.danger : C.accent} sub="Past due date" />
          <StatCard label="Due this week" value={String(myOpen.filter((t) => t.dueDate && !taskIsOverdue(t) && fromKey(t.dueDate) <= addDays(TODAY, 7)).length)} sub="Next 7 days" />
        </div>
      )}

      <Panel style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
          {mode === "all" && (
            <Field label="Assigned to">
              <select style={{ ...inputStyle, width: 180 }} value={assigneeFilter} onChange={(e) => setAssigneeFilter(e.target.value)}>
                <option value="all">Everyone</option>
                <option value="unassigned">Unassigned</option>
                {employees.filter((e) => e.active).map((e) => <option key={e.id} value={e.id}>{e.name}</option>)}
              </select>
            </Field>
          )}
          <Field label="Client">
            <select style={{ ...inputStyle, width: 190 }} value={clientFilter} onChange={(e) => setClientFilter(e.target.value)}>
              <option value="all">All clients</option>
              {clients.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
          <Field label="Status">
            <select style={{ ...inputStyle, width: 160 }} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
              <option value="open">Open (not done)</option>
              <option value="overdue">Overdue</option>
              <option value="all">All</option>
              {Object.entries(TASK_STATUSES).map(([k, l]) => <option key={k} value={k}>{l}</option>)}
            </select>
          </Field>
        </div>
        <TaskList user={user} tasks={visible} clients={clients} employees={employees} onEdit={(t) => setTaskModal(t)} onQuickStatus={quickStatus} onOpenCard={onOpenCard} />
      </Panel>

      {mode === "mine" && recentNotes.length > 0 && (
        <Panel>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 10 }}>My recent client notes</div>
          {recentNotes.map((n) => (
            <div key={n.id} style={{ display: "flex", gap: 10, padding: "6px 0", borderBottom: `1px solid ${C.border}`, fontFamily: sans, fontSize: 12.5 }}>
              <span style={{ fontFamily: mono, color: C.inkMuted, minWidth: 80 }}>{n.noteDate}</span>
              <Badge tone="neutral">{NOTE_KINDS[n.kind]}</Badge>
              <span style={{ fontWeight: 600, minWidth: 140 }}>{clients.find((c) => c.id === n.clientId)?.name || "—"}</span>
              <span style={{ color: C.ink, flex: 1 }}>{n.content}</span>
            </div>
          ))}
        </Panel>
      )}

      {taskModal && <TaskModal user={user} task={taskModal === "new" ? null : taskModal} clients={clients} employees={employees} onClose={() => setTaskModal(null)} onSaved={refetchTasks} />}
      {noteModal && <NoteModal user={user} clients={clients} onClose={() => setNoteModal(false)} onSaved={refetchNotes} />}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Client card — Fees & hours tab (the single place for billing terms)     */
/* ---------------------------------------------------------------------- */

// One-off / yearly service fee on the client card (annual accounts + income
// tax return, or start-up & business formation). "Applicable" switches it on;
// "Post to revenue" creates the matching extra fee so it flows into revenue,
// profitability and the billing monitor. Saving afterwards keeps that extra
// fee in sync; switching it off offers to remove it.
function OneOffServiceCard({ client, kind, title, hint, notePrefix, extras, refetchClients }) {
  const K = (suffix) => client[`${kind}${suffix}`];
  const col = (c) => `${kind}_fee_${c}`;
  const [enabled, setEnabled] = useState(!!K("FeeEnabled"));
  const [amount, setAmount] = useState(K("Fee") ?? "");
  const [month, setMonth] = useState(K("FeeMonth") || toKey(TODAY).slice(0, 7));
  const [cc, setCc] = useState(K("FeeCc"));
  const [busy, setBusy] = useState(false);
  useEffect(() => {
    setEnabled(!!client[`${kind}FeeEnabled`]); setAmount(client[`${kind}Fee`] ?? ""); setMonth(client[`${kind}FeeMonth`] || toKey(TODAY).slice(0, 7)); setCc(client[`${kind}FeeCc`]);
  }, [client.id, client[`${kind}FeeEnabled`], client[`${kind}Fee`], client[`${kind}FeeMonth`], client[`${kind}FeeCc`], client[`${kind}FeeExtraId`]]); // eslint-disable-line react-hooks/exhaustive-deps
  const linked = extras.find((x) => x.id === K("FeeExtraId")) || null;
  const note = `${notePrefix} — ${client.name}`;
  async function save() {
    if (enabled && (!amount || Number(amount) <= 0 || !month)) { window.alert("Enter the amount (greater than zero) and the month."); return; }
    setBusy(true);
    try {
      let extraId = K("FeeExtraId");
      if (linked && enabled) {
        const { error } = await supabase.from("client_extra_fees").update({ month, amount: Number(amount), cost_center: cc, note }).eq("id", linked.id);
        if (error) { window.alert(`Couldn't update the revenue entry: ${error.message}`); return; }
      }
      if (linked && !enabled) {
        if (!window.confirm(`Also remove the ${fmtEur(linked.amount)} entry (${linked.month}) from revenue?`)) return;
        await supabase.from("client_extra_fees").delete().eq("id", linked.id);
        extraId = null;
      }
      const { error } = await supabase.from("clients").update({ [col("enabled")]: enabled, [`${kind}_fee`]: amount ? Number(amount) : null, [col("month")]: month || null, [col("cc")]: cc, [col("extra_id")]: extraId }).eq("id", client.id);
      if (error) { window.alert(`Couldn't save: ${error.message}`); return; }
      await refetchClients();
    } finally { setBusy(false); }
  }
  async function post() {
    if (!enabled || !amount || Number(amount) <= 0 || !month) { window.alert("Switch it on and enter the amount and month first."); return; }
    setBusy(true);
    try {
      const { data, error } = await supabase.from("client_extra_fees").insert({ client_id: client.id, month, amount: Number(amount), cost_center: cc, note }).select().single();
      if (error || !data) { window.alert(`Couldn't post to revenue: ${error ? error.message : "unknown error"}`); return; }
      const { error: e2 } = await supabase.from("clients").update({ [col("enabled")]: true, [`${kind}_fee`]: Number(amount), [col("month")]: month, [col("cc")]: cc, [col("extra_id")]: data.id }).eq("id", client.id);
      if (e2) { window.alert(`Posted, but couldn't link it: ${e2.message}`); }
      await refetchClients();
    } finally { setBusy(false); }
  }
  return (
    <div style={{ border: `1px solid ${enabled ? C.accent : C.border}`, borderRadius: 8, padding: 14, background: enabled ? C.accentSoft : undefined }}>
      <label style={{ display: "flex", alignItems: "center", gap: 8, fontFamily: sans, fontWeight: 700, fontSize: 13.5, color: C.ink, cursor: "pointer" }}>
        <input type="checkbox" checked={enabled} onChange={(e) => setEnabled(e.target.checked)} />
        {title} — applicable
      </label>
      <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, margin: "4px 0 10px" }}>{hint}</div>
      {enabled && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
          <Field label="Fee (€)"><input type="number" min="0" style={inputStyle} value={amount} onChange={(e) => setAmount(e.target.value)} /></Field>
          <Field label="Month billed"><input type="month" style={inputStyle} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
          <Field label="Cost center">
            <select style={inputStyle} value={cc} onChange={(e) => setCc(e.target.value)}>{COST_CENTERS.map((c) => <option key={c} value={c}>{COST_CENTER_LABELS[c]}</option>)}</select>
          </Field>
        </div>
      )}
      <div style={{ display: "flex", gap: 8, alignItems: "center", marginTop: 10, flexWrap: "wrap" }}>
        <Btn size="sm" onClick={save} disabled={busy}>Save</Btn>
        {enabled && !linked && <Btn size="sm" variant="secondary" onClick={post} disabled={busy}>Post to revenue</Btn>}
        {linked && <span style={{ fontFamily: sans, fontSize: 12, color: C.accent, fontWeight: 600 }}>✓ In revenue: {fmtEur(linked.amount)} · {linked.month} · {COST_CENTER_LABELS[linked.costCenter] || linked.costCenter}</span>}
        {!linked && enabled && <span style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted }}>Not in revenue / billing yet</span>}
      </div>
    </div>
  );
}

function ClientFeesTab({ client, entries, employees, refetchClients, isAdmin = true }) {
  const [showRate, setShowRate] = useState(() => { try { return localStorage.getItem("gloria_show_rate") !== "0"; } catch { return true; } });
  const toggleRate = () => { const n = !showRate; setShowRate(n); try { localStorage.setItem("gloria_show_rate", n ? "1" : "0"); } catch { /* ignore */ } };
  const [viewPeriod, setViewPeriod] = useState("month");
  const [viewAnchor, setViewAnchor] = useState(TODAY);
  const cur = currentFeeOf(client);
  const blankForm = () => ({
    fixedFee: client.fixedFee ?? "", effectiveDate: toKey(TODAY), endDate: client.endDate || "",
    split: Object.fromEntries(COST_CENTERS.map((cc) => [cc, cur.feeSplit?.[cc] ? String(cur.feeSplit[cc]) : ""])), splitNote: cur.feeSplitNote || "",
    alloc: Object.fromEntries(CATEGORIES.map((c) => [c, cur.allocation?.[c] ? String(cur.allocation[c]) : ""])),
  });
  const [form, setForm] = useState(blankForm);
  const [editingHistId, setEditingHistId] = useState(null);
  useEffect(() => { setForm(blankForm()); setEditingHistId(null); }, [client.id, client.fixedFee, client.feeSplitNote]); // eslint-disable-line react-hooks/exhaustive-deps
  const [xf, setXf] = useState({ month: toKey(TODAY).slice(0, 7), amount: "", costCenter: "accounting", note: "" });
  const [xfEditingId, setXfEditingId] = useState(null);

  const fee = form.fixedFee === "" ? null : Number(form.fixedFee);
  const split = Object.fromEntries(COST_CENTERS.map((cc) => [cc, Number(form.split[cc]) || 0]));
  const splitTotal = feeSplitTotal(split);
  const remaining = (fee || 0) - splitTotal;
  const setSplit = (cc, v) => setForm({ ...form, split: { ...form.split, [cc]: v } });
  const setAlloc = (c, v) => setForm({ ...form, alloc: { ...form.alloc, [c]: v } });
  const feeRow = () => ({
    fixed_fee: fee,
    alloc_accounting: Number(form.alloc.accounting) || 0, alloc_tax: Number(form.alloc.tax) || 0, alloc_payroll: Number(form.alloc.payroll) || 0, alloc_other: Number(form.alloc.other) || 0,
    ...Object.fromEntries(COST_CENTERS.map((cc) => [`cc_${cc}`, split[cc] || 0])), cc_other_note: split.other > 0 ? form.splitNote.trim() : null,
  });
  function validate() {
    if (fee && splitTotal > 0 && Math.abs(splitTotal - fee) > 0.01) { window.alert(`The cost-center split adds up to ${fmtEur(splitTotal)}, but the fee is ${fmtEur(fee)}.`); return false; }
    if (!fee && splitTotal > 0) { window.alert("Enter the fixed fee first."); return false; }
    if (split.other > 0 && !form.splitNote.trim()) { window.alert('A note is required for the "Other" part of the fee.'); return false; }
    return true;
  }
  async function saveTerms() {
    if (!validate()) return;
    const row = { ...feeRow(), client_id: client.id, effective_date: form.effectiveDate || toKey(TODAY) };
    const { error } = editingHistId
      ? await supabase.from("client_fee_history").update(row).eq("id", editingHistId)
      : await supabase.from("client_fee_history").insert(row);
    if (error) { window.alert(`Couldn't save: ${error.message}`); return; }
    const { error: e2 } = await supabase.from("clients").update({ fixed_fee: fee, end_date: form.endDate || null }).eq("id", client.id);
    if (e2) { window.alert(`Couldn't save: ${e2.message}`); return; }
    await refetchClients();
    setEditingHistId(null);
  }
  function editHistory(h) {
    setEditingHistId(h.id);
    setForm({ ...form, fixedFee: h.fixedFee ?? "", effectiveDate: h.effectiveDate, split: Object.fromEntries(COST_CENTERS.map((cc) => [cc, h.feeSplit?.[cc] ? String(h.feeSplit[cc]) : ""])), splitNote: h.feeSplitNote || "", alloc: Object.fromEntries(CATEGORIES.map((c) => [c, h.allocation?.[c] ? String(h.allocation[c]) : ""])) });
  }
  async function deleteHistory(h) {
    if (!window.confirm(`Delete the fee terms effective ${h.effectiveDate}?`)) return;
    await supabase.from("client_fee_history").delete().eq("id", h.id);
    refetchClients();
  }
  async function saveExtraFee() {
    if (!xf.month || !xf.amount || Number(xf.amount) <= 0) { window.alert("Enter the month and an amount greater than zero."); return; }
    if (xf.costCenter === "other" && !xf.note.trim()) { window.alert('A note is required when the cost center is "Other".'); return; }
    const row = { month: xf.month, amount: Number(xf.amount), note: xf.note.trim() || null, cost_center: xf.costCenter };
    const { error } = xfEditingId
      ? await supabase.from("client_extra_fees").update(row).eq("id", xfEditingId)
      : await supabase.from("client_extra_fees").insert({ client_id: client.id, ...row });
    if (error) { window.alert(`Couldn't save extra fee: ${error.message}`); return; }
    await refetchClients();
    setXf({ month: xf.month, amount: "", costCenter: xf.costCenter, note: "" }); setXfEditingId(null);
  }
  async function deleteExtraFee(x) {
    if (!window.confirm(`Delete the ${fmtEur(x.amount)} extra fee of ${x.month}?`)) return;
    await supabase.from("client_extra_fees").delete().eq("id", x.id);
    refetchClients();
  }

  const history = [...(client.feeHistory || [])].sort((a, b) => b.effectiveDate.localeCompare(a.effectiveDate));
  const extras = [...(client.extraFeeEntries || [])].sort((a, b) => b.month.localeCompare(a.month));
  const [rStart, rEnd] = periodRange(viewPeriod, viewAnchor);
  const periodEntries = entries.filter((e) => e.clientId === client.id && fromKey(e.date) >= rStart && fromKey(e.date) <= rEnd);
  const metrics = periodClientMetrics(client, periodEntries, rStart, rEnd);
  const hoursByEmployee = employees.map((emp) => ({ emp, hours: periodEntries.filter((e) => e.employeeId === emp.id).reduce((s, e) => s + e.hours, 0) })).filter((x) => x.hours > 0).sort((a, b) => b.hours - a.hours);
  const totalHours = periodEntries.reduce((s, e) => s + e.hours, 0);
  const empName = (id) => employees.find((e) => e.id === id)?.name || "—";

  // Effective rate by month (last 12 months): revenue ÷ hours worked for this client.
  // Cost = each employee's monthly cost × their share of hours spent on this client that month.
  const rateRows = Array.from({ length: 12 }, (_, i) => {
    const ms = new Date(TODAY.getFullYear(), TODAY.getMonth() - i, 1);
    const me = endOfMonth(ms);
    const inMonth = (e) => { const d = fromKey(e.date); return d >= ms && d <= me; };
    const monthEntries = entries.filter(inMonth);
    const cEntries = monthEntries.filter((e) => e.clientId === client.id);
    const hours = cEntries.reduce((s, e) => s + e.hours, 0);
    const revenue = periodClientMetrics(client, cEntries, ms, me).revenue || 0;
    let cost = 0;
    employees.forEach((emp) => {
      const mine = cEntries.filter((e) => e.employeeId === emp.id).reduce((s, e) => s + e.hours, 0);
      if (!mine) return;
      const all = monthEntries.filter((e) => e.employeeId === emp.id).reduce((s, e) => s + e.hours, 0);
      if (all > 0) cost += periodEmployeeCost(emp, ms, me) * (mine / all);
    });
    return { key: toKey(ms).slice(0, 7), hours, revenue, cost, rate: hours > 0 ? revenue / hours : null, costH: hours > 0 ? cost / hours : null };
  }).filter((r) => r.hours > 0 || r.revenue > 0);
  const rateTot = rateRows.reduce((a, r) => ({ hours: a.hours + r.hours, revenue: a.revenue + r.revenue, cost: a.cost + r.cost }), { hours: 0, revenue: 0, cost: 0 });
  const fmtRate = (n) => (n == null ? "—" : "€" + n.toFixed(2));

  return (
    <div>
      {isAdmin && <Panel style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink }}>Effective rate by month</div>
          <Btn size="sm" variant="ghost" onClick={toggleRate}>{showRate ? "Hide" : "Open"}</Btn>
        </div>
        {showRate && <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, margin: "2px 0 10px" }}>Effective rate = revenue ÷ hours worked for this client. Compare it with cost/h: if the rate is below cost/h, the client loses money. Last 12 months.</div>}
        {!showRate ? null : rateRows.length === 0 ? <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkFaint }}>No hours or revenue yet.</div> : (
          <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: sans, fontSize: 12.5 }}>
            <thead><tr>{["Month", "Hours", "Revenue", "Effective rate", "Cost", "Cost /h", "Profit"].map((h, i) => <th key={h} style={{ textAlign: i ? "right" : "left", padding: "6px 8px", color: C.inkMuted, fontWeight: 600, borderBottom: `1px solid ${C.border || "#ddd"}` }}>{h}</th>)}</tr></thead>
            <tbody>
              {[...rateRows, { key: "Total", ...rateTot, rate: rateTot.hours > 0 ? rateTot.revenue / rateTot.hours : null, costH: rateTot.hours > 0 ? rateTot.cost / rateTot.hours : null, total: true }].map((r) => {
                const profit = r.revenue - r.cost;
                return (
                  <tr key={r.key} style={{ fontWeight: r.total ? 700 : 400 }}>
                    <td style={{ padding: "6px 8px" }}>{r.key}</td>
                    <td style={{ padding: "6px 8px", textAlign: "right", fontFamily: mono }}>{r.hours.toFixed(1)}</td>
                    <td style={{ padding: "6px 8px", textAlign: "right", fontFamily: mono }}>{fmtEur(r.revenue)}</td>
                    <td style={{ padding: "6px 8px", textAlign: "right", fontFamily: mono, color: r.rate != null && r.costH != null && r.rate < r.costH ? C.danger : C.ink }}>{fmtRate(r.rate)}</td>
                    <td style={{ padding: "6px 8px", textAlign: "right", fontFamily: mono }}>{fmtEur(r.cost)}</td>
                    <td style={{ padding: "6px 8px", textAlign: "right", fontFamily: mono }}>{fmtRate(r.costH)}</td>
                    <td style={{ padding: "6px 8px", textAlign: "right", fontFamily: mono, color: profit < 0 ? C.danger : C.accent }}>{fmtEur(profit)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </Panel>}
      <Panel style={{ marginBottom: 18 }}>
        <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>Annual & start-up services</div>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Switch on only what applies to this client. "Post to revenue" adds the fee as an extra fee, so it appears in revenue, profitability and the billing monitor.</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
          <OneOffServiceCard client={client} kind="annual" title="Annual fee" hint="Annual financial statements (ισολογισμός) and income tax return (δήλωση φορολογίας εισοδήματος)." notePrefix="Annual fee: financial statements & income tax return" extras={client.extraFeeEntries || []} refetchClients={refetchClients} />
          <OneOffServiceCard client={client} kind="formation" title="Start-up & business formation" hint="Start-up work (έναρξη) and business formation: ΥΜΣ, registrations, bank, ΕΦΚΑ, ΓΕΜΗ." notePrefix="Start-up & business formation" extras={client.extraFeeEntries || []} refetchClients={refetchClients} />
        </div>
      </Panel>
      <div style={{ display: "grid", gridTemplateColumns: "3fr 2fr", gap: 18, alignItems: "start" }}>
        <Panel>
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 2 }}>{editingHistId ? "Edit fee terms" : "Fee terms"}</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 12 }}>Saving with different numbers adds a new row to the history, effective from the date below — past periods are never overwritten.</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
            <Field label="Fixed monthly fee (€, net)"><input type="number" min="0" style={inputStyle} value={form.fixedFee} onChange={(e) => setForm({ ...form, fixedFee: e.target.value })} /></Field>
            <Field label="Effective from"><input type="date" style={inputStyle} value={form.effectiveDate} onChange={(e) => setForm({ ...form, effectiveDate: e.target.value })} /></Field>
            <Field label="Client end date (optional)"><input type="date" style={inputStyle} value={form.endDate} onChange={(e) => setForm({ ...form, endDate: e.target.value })} /></Field>
          </div>
          <div style={{ display: "flex", alignItems: "flex-end", gap: 14, marginTop: 10, flexWrap: "wrap" }}>
            <Field label="ΦΠΑ on fees">
              <select style={{ ...inputStyle, width: 170 }} value={String(client.vatRate ?? 24)}
                onChange={async (e) => { const { error } = await supabase.from("clients").update({ vat_rate: Number(e.target.value) }).eq("id", client.id); if (error) { window.alert(`Couldn't save: ${error.message}`); return; } refetchClients(); }}>
                <option value="24">24%</option>
                <option value="0">0% (no VAT)</option>
              </select>
            </Field>
            <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted, paddingBottom: 8 }}>
              {fee ? <>Payable per month: <strong style={{ fontFamily: mono, color: C.ink }}>{fmtEur(fee * (1 + (client.vatRate ?? 24) / 100))}</strong> ({fmtEur(fee)} + ΦΠΑ {fmtEur(fee * (client.vatRate ?? 24) / 100)})</> : "Revenue and profitability stay net of VAT."}
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "flex-end", gap: 14, marginTop: 10, flexWrap: "wrap" }}>
            {[
              ["invoice_type", "Invoice type (myDATA)", String(client.invoiceType || "2.1"), Object.entries(MYDATA_INVOICE_TYPES).map(([k, v]) => [k, `${k} — ${v.label}`]), 300],
              ["fixed_service", "Service on the monthly invoice", client.fixedService || "accounting", Object.entries(INVOICE_SERVICES), 220],
            ].map(([col, label, val, opts, w]) => (
              <Field key={col} label={label}>
                <select style={{ ...inputStyle, width: w }} value={val}
                  onChange={async (e) => { const { error } = await supabase.from("clients").update({ [col]: e.target.value }).eq("id", client.id); if (error) { window.alert(`Couldn't save: ${error.message}`); return; } refetchClients(); }}>
                  {opts.map(([k, l]) => <option key={k} value={k}>{l}</option>)}
                </select>
              </Field>
            ))}
            {Number(client.vatRate ?? 24) === 0 && (
              <Field label="VAT exemption reason (0% VAT)">
                <select style={{ ...inputStyle, width: 300 }} value={client.vatExemptionCategory ?? ""}
                  onChange={async (e) => { const { error } = await supabase.from("clients").update({ vat_exemption_category: e.target.value === "" ? null : Number(e.target.value) }).eq("id", client.id); if (error) { window.alert(`Couldn't save: ${error.message}`); return; } refetchClients(); }}>
                  <option value="">— choose —</option>
                  {MYDATA_VAT_EXEMPTIONS.map(([k, l]) => <option key={k} value={k}>{k} — {l}</option>)}
                </select>
              </Field>
            )}
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", margin: "12px 0 6px" }}>
            <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink }}>Split of the fee by cost center (€/month)</div>
            <div style={{ fontFamily: mono, fontSize: 11.5, color: splitTotal === 0 ? C.inkFaint : Math.abs(remaining) <= 0.01 ? C.accent : C.danger }}>
              {splitTotal === 0 ? "No split yet" : Math.abs(remaining) <= 0.01 ? "Matches fee ✓" : remaining > 0 ? `${fmtEur(remaining)} left` : `${fmtEur(-remaining)} over`}
            </div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8 }}>
            {COST_CENTERS.map((cc) => <Field key={cc} label={COST_CENTER_LABELS[cc]}><input type="number" min="0" style={inputStyle} value={form.split[cc]} onChange={(e) => setSplit(cc, e.target.value)} placeholder="0" /></Field>)}
          </div>
          {Number(form.split.other) > 0 && <div style={{ marginTop: 8 }}><Field label='Note — what does "Other" cover? (required)'><input style={inputStyle} value={form.splitNote} onChange={(e) => setForm({ ...form, splitNote: e.target.value })} /></Field></div>}
          <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink, margin: "12px 0 6px" }}>Agreed monthly hours by category (workload only)</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8 }}>
            {CATEGORIES.map((c) => <Field key={c} label={CATEGORY_LABELS[c]}><input type="number" min="0" style={inputStyle} value={form.alloc[c]} onChange={(e) => setAlloc(c, e.target.value)} placeholder="h" /></Field>)}
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
            <Btn onClick={saveTerms}>{editingHistId ? "Save changes to this row" : "Save fee terms"}</Btn>
            {editingHistId && <Btn variant="secondary" onClick={() => { setEditingHistId(null); setForm(blankForm()); }}>Cancel edit</Btn>}
          </div>

          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 13.5, color: C.ink, margin: "20px 0 8px" }}>Fee history</div>
          {history.length === 0 ? <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkFaint }}>No fee terms recorded yet.</div> : (
            <TableShell headers={["From", "Fixed fee", "Split", "Hours alloc.", ""]}>
              {history.map((h, i) => {
                const to = i === 0 ? "now" : addDays(fromKey(history[i - 1].effectiveDate), -1);
                const sp = h.feeSplit || {}; const spTotal = feeSplitTotal(sp);
                return (
                  <tr key={h.id} style={{ background: editingHistId === h.id ? C.accentSoft : undefined }}>
                    <Td mono>{h.effectiveDate} → {typeof to === "string" ? to : toKey(to)}</Td>
                    <Td mono>{h.fixedFee ? fmtEur(h.fixedFee) : "—"}</Td>
                    <Td style={{ fontSize: 12, color: spTotal > 0 ? C.ink : C.warn }} title={COST_CENTERS.filter((cc) => sp[cc]).map((cc) => `${COST_CENTER_LABELS[cc]} ${fmtEur(sp[cc])}`).join(" · ")}>{spTotal > 0 ? COST_CENTERS.filter((cc) => sp[cc]).map((cc) => `${COST_CENTER_LABELS[cc].slice(0, 3)} ${sp[cc]}`).join(" · ") : (h.fixedFee ? "Not split" : "—")}</Td>
                    <Td mono>{CATEGORIES.reduce((s, c) => s + (h.allocation?.[c] || 0), 0)}h</Td>
                    <Td right><div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}><button onClick={() => editHistory(h)} style={iconBtnStyle}><Pencil size={12} color={C.inkMuted} /></button>{history.length > 1 && <button onClick={() => deleteHistory(h)} style={iconBtnStyle}><Trash2 size={12} color={C.inkMuted} /></button>}</div></Td>
                  </tr>
                );
              })}
            </TableShell>
          )}
        </Panel>

        <div>
          <Panel style={{ marginBottom: 18 }}>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 8 }}>{xfEditingId ? "Edit extra fee" : "Add extra fee"}</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              <Field label="Month"><input type="month" style={inputStyle} value={xf.month} onChange={(e) => setXf({ ...xf, month: e.target.value })} /></Field>
              <Field label="Amount (€)"><input type="number" min="0" style={inputStyle} value={xf.amount} onChange={(e) => setXf({ ...xf, amount: e.target.value })} placeholder="e.g. 150" /></Field>
              <Field label="Cost center">
                <select style={inputStyle} value={xf.costCenter} onChange={(e) => setXf({ ...xf, costCenter: e.target.value })}>{COST_CENTERS.map((cc) => <option key={cc} value={cc}>{COST_CENTER_LABELS[cc]}</option>)}</select>
              </Field>
              <Field label={xf.costCenter === "other" ? "Note (required)" : "Note"}><input style={inputStyle} value={xf.note} onChange={(e) => setXf({ ...xf, note: e.target.value })} placeholder="e.g. Extra VAT filing" /></Field>
            </div>
            <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
              <Btn size="sm" icon={Plus} onClick={saveExtraFee}>{xfEditingId ? "Save changes" : "Add extra fee"}</Btn>
              {xfEditingId && <Btn size="sm" variant="secondary" onClick={() => { setXfEditingId(null); setXf({ ...xf, amount: "", note: "" }); }}>Cancel</Btn>}
            </div>
            {extras.length > 0 && (
              <div style={{ marginTop: 12 }}>
                <TableShell headers={["Month", "Amount", "Cost center", "Note", ""]}>
                  {extras.map((x) => (
                    <tr key={x.id}>
                      <Td mono>{x.month}</Td>
                      <Td mono>{fmtEur(x.amount)}</Td>
                      <Td style={{ fontSize: 12 }}>{costCenterLabel(x.costCenter)}</Td>
                      <Td style={{ color: C.inkMuted, fontSize: 12 }}>{x.note || "—"}</Td>
                      <Td right><div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}><button onClick={() => { setXfEditingId(x.id); setXf({ month: x.month, amount: String(x.amount), costCenter: x.costCenter, note: x.note || "" }); }} style={iconBtnStyle}><Pencil size={12} color={C.inkMuted} /></button><button onClick={() => deleteExtraFee(x)} style={iconBtnStyle}><Trash2 size={12} color={C.inkMuted} /></button></div></Td>
                    </tr>
                  ))}
                </TableShell>
              </div>
            )}
          </Panel>

          <Panel>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 8 }}>Hours & revenue</div>
            <PeriodSelector period={viewPeriod} setPeriod={setViewPeriod} anchor={viewAnchor} setAnchor={setViewAnchor} />
            <div style={{ display: "flex", gap: 14, margin: "12px 0", flexWrap: "wrap" }}>
              {isAdmin && <StatCard label="Hours" value={totalHours.toFixed(2) + "h"} sub={metrics.allocation ? `of ${metrics.allocation.toFixed(1)}h allocated` : "No allocation set"} />}
              <StatCard label="Revenue" value={fmtEur(metrics.revenue)} accent={C.accent} sub="Fixed + extra fees, accrued" />
            </div>
            {!isAdmin ? null : hoursByEmployee.length === 0 ? <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkFaint }}>No hours logged in this period.</div> : (
              <TableShell headers={["Employee", "Hours", ...CATEGORIES.map((c) => CATEGORY_LABELS[c])]}>
                {hoursByEmployee.map(({ emp, hours }) => (
                  <tr key={emp.id}>
                    <Td>{emp.name}</Td>
                    <Td mono style={{ fontWeight: 700 }}>{hours.toFixed(2)}h</Td>
                    {CATEGORIES.map((c) => <Td key={c} mono style={{ color: C.inkMuted }}>{periodEntries.filter((e) => e.employeeId === emp.id && e.category === c).reduce((s, e) => s + e.hours, 0).toFixed(1)}</Td>)}
                  </tr>
                ))}
              </TableShell>
            )}
          </Panel>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Client card — Documents tab (Supabase Storage uploads)                  */
/* ---------------------------------------------------------------------- */

const DOCUMENT_CATEGORIES = [
  { key: "contract", label: "Signed agreement / engagement letter" },
  ...AML_DOCUMENTS.map((d) => ({ key: d.key, label: `AML — ${d.label}` })),
  { key: "authorisation", label: "Authorisation (ΑΑΔΕ / ΕΦΚΑ / ΕΡΓΑΝΗ)" },
  { key: "other", label: "Other" },
];
const documentCategoryLabel = (k) => DOCUMENT_CATEGORIES.find((c) => c.key === k)?.label || k;

async function uploadClientDocument(client, file, category, note, user) {
  const dot = file.name.lastIndexOf(".");
  const ext = dot > 0 ? file.name.slice(dot).replace(/[^A-Za-z0-9.]/g, "").slice(0, 10) : "";
  const base = (dot > 0 ? file.name.slice(0, dot) : file.name).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Za-z0-9._-]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 60) || "file";
  const safe = base + ext;
  const path = `${client.id}/${Date.now()}_${safe}`;
  const { error } = await supabase.storage.from(DOCUMENT_BUCKET).upload(path, file, { upsert: false, contentType: file.type || undefined });
  if (error) { window.alert(`Upload failed: ${error.message}\n\nIf this says the bucket does not exist or access is denied, run the latest schema.sql in Supabase (it creates the "client-documents" bucket and its policies).`); return false; }
  const { error: e2 } = await supabase.from("client_documents").insert({ client_id: client.id, category, file_name: file.name, file_path: path, file_size: file.size, mime_type: file.type || null, note: note || null, uploaded_by: user.id });
  if (e2) { window.alert(`Upload saved but could not be recorded: ${e2.message}`); return false; }
  return true;
}
async function openClientDocument(doc) {
  const { data, error } = await supabase.storage.from(DOCUMENT_BUCKET).createSignedUrl(doc.filePath, 600);
  if (error || !data?.signedUrl) { window.alert(`Couldn't open the file: ${error?.message || "no URL"}`); return; }
  window.open(data.signedUrl, "_blank");
}
async function deleteClientDocument(doc) {
  if (!window.confirm(`Delete "${doc.fileName}"? This removes the file permanently.`)) return false;
  await supabase.storage.from(DOCUMENT_BUCKET).remove([doc.filePath]);
  const { error } = await supabase.from("client_documents").delete().eq("id", doc.id);
  if (error) { window.alert(`Couldn't delete: ${error.message}`); return false; }
  return true;
}

// Small upload control: a hidden file input behind a button.
function UploadButton({ label = "Upload", size = "sm", onFile }) {
  const ref = React.useRef(null);
  const [busy, setBusy] = useState(false);
  return (
    <>
      <input ref={ref} type="file" style={{ display: "none" }} onChange={async (e) => { const f = e.target.files?.[0]; e.target.value = ""; if (!f) return; setBusy(true); try { await onFile(f); } finally { setBusy(false); } }} />
      <Btn size={size} variant="secondary" onClick={() => ref.current?.click()} disabled={busy}>{busy ? "Uploading…" : label}</Btn>
    </>
  );
}

// Εξουσιοδότηση προς λογιστή-φοροτεχνικό για ΑΑΔΕ / myDATA — pre-filled from the client card; signed by the client's representative.
function buildAuthorisationHtml({ client, aml, firm = DEFAULT_FIRM, contacts = [] }) {
  const e = escapeHtml;
  const a = aml || {};
  const natural = (a.entityType || amlEntityTypeFor(client.legalForm)) === "natural";
  const b = (x, ph) => (x ? `<strong>${e(x)}</strong>` : `<span style="color:#999">${ph}</span>`);
  const primary = contacts.find((c) => c.isPrimary) || contacts[0] || {};
  const repName = a.representativeName || client.representativeName || (natural ? client.name : "");
  const repId = a.representativeIdNumber || client.representativeIdNumber || "";
  const repAfm = a.representativeAfm || client.representativeAfm || (natural ? client.afm : "");
  const email = client.representativeEmail || primary.email || client.email || "";
  const phone = client.representativePhone || primary.phone || client.phone || "";
  const who = natural ? "τον/την" : "τον/την";
  const services = [
    "Να αποκτήσουν και να διατηρούν πρόσβαση στις ηλεκτρονικές υπηρεσίες της Ανεξάρτητης Αρχής Δημοσίων Εσόδων (myAADE / TAXISnet), τόσο με χρήση των κωδικών πρόσβασης " + (natural ? "μου" : "της εταιρείας") + " όσο και μέσω της διαδικασίας ορισμού λογιστή – εκπροσώπου και της υποβολής των σχετικών εξουσιοδοτήσεων χρηστών.",
    "Να συντάσσουν, να υπογράφουν και να υποβάλλουν ηλεκτρονικά κάθε είδους φορολογική δήλωση και κατάσταση " + (natural ? "μου" : "της εταιρείας") + ", ενδεικτικά: δηλώσεις Φ.Π.Α., δηλώσεις παρακρατούμενων και προκαταβλητέων φόρων, δήλωση φορολογίας εισοδήματος " + (natural ? "φυσικών προσώπων (Ε1, Ε2, Ε3 κ.λπ.)" : "νομικών προσώπων με τα συνυποβαλλόμενα έντυπα (Ν, Ε3 κ.λπ.)") + ", ανακεφαλαιωτικούς πίνακες ενδοκοινοτικών συναλλαγών (VIES, Intrastat), καθώς και τυχόν τροποποιητικές αυτών.",
    "Να διαβιβάζουν δεδομένα στην ψηφιακή πλατφόρμα myDATA, να διαχειρίζονται τα σχετικά κλειδιά και τους παρόχους ηλεκτρονικής τιμολόγησης και να προβαίνουν σε κάθε αναγκαία ενέργεια χαρακτηρισμού και συμφωνίας των δεδομένων.",
    "Να υποβάλλουν δηλώσεις μεταβολής στοιχείων μητρώου, έναρξης ή διακοπής δραστηριοτήτων, μεταβολής έδρας, εκπροσώπων και λοιπών στοιχείων " + (natural ? "μου" : "της εταιρείας") + ", πάντοτε κατόπιν προηγούμενης έγκρισης και ενημέρωσης " + (natural ? "δική μου" : "του νόμιμου εκπροσώπου") + ".",
    "Να αιτούνται και να παραλαμβάνουν βεβαιώσεις, πιστοποιητικά, αποδεικτικά φορολογικής ενημερότητας, βεβαιώσεις τρέχουσας εικόνας οντότητας, εκτυπώσεις εκκαθαριστικών σημειωμάτων, ταυτότητες οφειλής και κάθε άλλο έγγραφο εκδιδόμενο από την Α.Α.Δ.Ε.",
    "Να έχουν πρόσβαση στην προσωποποιημένη πληροφόρηση και στη θυρίδα ηλεκτρονικής κοινοποίησης εγγράφων " + (natural ? "μου" : "της εταιρείας") + ", να λαμβάνουν γνώση των κοινοποιούμενων πράξεων και να υποβάλλουν αιτήσεις, υπομνήματα και υπεύθυνες δηλώσεις, πάντοτε κατόπιν προηγούμενης έγκρισης και ενημέρωσης " + (natural ? "δική μου" : "του νόμιμου εκπροσώπου") + ".",
    "Να προβαίνουν σε ρυθμίσεις οφειλών και να παρακολουθούν την εικόνα οφειλών " + (natural ? "μου" : "της εταιρείας") + ".",
    (natural ? "Να με εκπροσωπούν" : "Να εκπροσωπούν την εταιρεία") + " ενώπιον της αρμόδιας Δ.Ο.Υ." + (client.doy ? " " + client.doy : "") + ", των Ελεγκτικών Κέντρων και κάθε άλλης υπηρεσίας της Α.Α.Δ.Ε., καταθέτοντας και παραλαμβάνοντας έγγραφα.",
  ];
  return `
  <h1 style="text-align:center">ΕΞΟΥΣΙΟΔΟΤΗΣΗ</h1>
  <p>${natural ? "Ο/Η κάτωθι υπογεγραμμένος/-η" : "Ο/Η κάτωθι υπογεγραμμένος/-η"} ${b(repName, "[ΟΝΟΜΑΤΕΠΩΝΥΜΟ ΚΑΙ ΠΑΤΡΩΝΥΜΟ]")}, με ΑΔΤ ${b(repId, ".....................")} και ΑΦΜ ${b(repAfm, "[ΑΦΜ]")}, κάτοικος ${b(natural ? client.address : "", "............................................")}, τηλ. ${b(phone, "..................")}, e-mail ${b(email, "..................")},</p>
  ${natural ? `<p>${client.legalForm === "Ιδιώτης" ? "" : `υπό την ιδιότητά μου ως ατομικού επιχειρηματία με ΑΦΜ ${b(client.afm, "[ΑΦΜ]")}, Δ.Ο.Υ. ${b(client.doy, "[ΔΟΥ]")}, με έδρα ${b(client.address, "[ΕΔΡΑ]")},`}</p>` : `<p>υπό την ιδιότητά μου ως <strong>διαχειριστής και νόμιμος εκπρόσωπος</strong> της εταιρείας με την επωνυμία «${b(client.name, "[ΕΠΩΝΥΜΙΑ]")}», με ΑΦΜ ${b(client.afm, "[ΑΦΜ]")}, Δ.Ο.Υ. ${b(client.doy, "[ΔΟΥ]")}, αριθμό Γ.Ε.ΜΗ. ${b(a.gemiNumber, "[ΓΕΜΗ]")}, με έδρα ${b(client.address, "[ΕΔΡΑ]")},</p>`}
  <p style="text-align:center"><strong>ΕΞΟΥΣΙΟΔΟΤΩ</strong></p>
  <p>τον <strong>${e(firm.representative)}</strong>${firm.representativeId ? `, με ΑΔΤ <strong>${e(firm.representativeId)}</strong>` : ""}${firm.representativeAfm ? ` και ΑΦΜ <strong>${e(firm.representativeAfm)}</strong>` : ""}, λογιστή – φοροτεχνικό, καθώς και την εταιρεία «<strong>${e(firm.name)}</strong>», με ΑΦΜ <strong>${e(firm.afm)}</strong>, να ενεργούν, από κοινού ή ο καθένας χωριστά, ως <strong>λογιστές – φοροτεχνικοί</strong> ${natural ? "μου" : "της ανωτέρω εταιρείας"} και ειδικότερα:</p>
  <ol>${services.map((s) => `<li>${e(s)}</li>`).join("")}</ol>
  <p><strong>${natural ? "Αναλαμβάνω" : "Ο νόμιμος εκπρόσωπος / διαχειριστής της εταιρείας υποχρεούται"} να παραδίδω έγκαιρα και πλήρως στους ανωτέρω λογιστές – φοροτεχνικούς όλα τα φορολογικά και λογιστικά στοιχεία, παραστατικά, τιμολόγια και κινήσεις (τραπεζικές και ταμειακές) ${natural ? "μου" : "της εταιρείας"}, ώστε να καθίσταται δυνατή η ορθή και εμπρόθεσμη εκπλήρωση των ανωτέρω υποχρεώσεων. Οι εξουσιοδοτούμενοι δεν φέρουν ευθύνη για παραλείψεις ή ανακρίβειες που οφείλονται στη μη έγκαιρη ή ελλιπή παράδοση των ως άνω στοιχείων.</strong></p>
  <p>Η παρούσα εξουσιοδότηση ισχύει από την ημερομηνία υπογραφής της και μέχρι την έγγραφη ανάκλησή της. Κάθε ενέργεια των ανωτέρω εξουσιοδοτούμενων στο πλαίσιο της παρούσας θεωρείται έγκυρη και δεσμευτική ${natural ? "για εμένα" : "για την εταιρεία"}.</p>
  <p style="margin-top:24px">......... / ......... / ${new Date().getFullYear()}</p>
  <p style="text-align:right;margin-top:30px"><strong>Ο εξουσιοδοτών</strong></p>
  <p style="text-align:right;margin-top:50px">...................................................<br/><strong>${e(repName || "")}</strong><br/><em class="muted">(θεώρηση γνησίου υπογραφής)</em></p>`;
}

function ClientDocumentsTab({ user, client, documents, employees, refetchDocuments, aml = null, ubos = [], contacts = [], firm = DEFAULT_FIRM }) {
  const [category, setCategory] = useState("other");
  const [note, setNote] = useState("");
  const [filter, setFilter] = useState("all");
  const empName = (id) => employees.find((e) => e.id === id)?.name || "—";
  const list = [...documents].filter((d) => filter === "all" || d.category === filter).sort((a, b) => b.uploadedAt.localeCompare(a.uploadedAt));
  return (
    <Panel>
      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 14, paddingBottom: 12, borderBottom: `1px solid ${C.border}` }}>
        <span style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted }}>Templates (pre-filled from the client card):</span>
        <Btn variant="secondary" icon={Receipt} onClick={() => printHtml(`Εξουσιοδότηση — ${client.name}`, buildAuthorisationHtml({ client, aml, firm, contacts }), { editable: true })}>Εξουσιοδότηση ΑΑΔΕ</Btn>
        <Btn variant="secondary" icon={Receipt} onClick={() => printHtml(`KYC — ${client.name}`, buildKycQuestionnaireHtml({ client, aml, ubos, firm }), { editable: true })}>KYC questionnaire</Btn>
      </div>
      <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap", marginBottom: 14 }}>
        <Field label="Category for the next upload">
          <select style={{ ...inputStyle, width: 320 }} value={category} onChange={(e) => setCategory(e.target.value)}>
            {DOCUMENT_CATEGORIES.map((c) => <option key={c.key} value={c.key}>{c.label}</option>)}
          </select>
        </Field>
        <Field label="Note (optional)"><input style={{ ...inputStyle, width: 240 }} value={note} onChange={(e) => setNote(e.target.value)} /></Field>
        <UploadButton size="md" label="Upload file" onFile={async (f) => { if (await uploadClientDocument(client, f, category, note, user)) { setNote(""); refetchDocuments(); } }} />
        <div style={{ marginLeft: "auto" }}>
          <Field label="Show">
            <select style={{ ...inputStyle, width: 220 }} value={filter} onChange={(e) => setFilter(e.target.value)}>
              <option value="all">All documents</option>
              {DOCUMENT_CATEGORIES.filter((c) => documents.some((d) => d.category === c.key)).map((c) => <option key={c.key} value={c.key}>{c.label}</option>)}
            </select>
          </Field>
        </div>
      </div>
      {list.length === 0 ? <EmptyState icon={Receipt} title="No documents yet" sub="Upload the engagement letter, identification documents and anything else that belongs to this client's file." /> : (
        <TableShell headers={["Uploaded", "Category", "File", "Size", "By", "Note", ""]}>
          {list.map((d) => (
            <tr key={d.id}>
              <Td mono>{d.uploadedAt.slice(0, 16).replace("T", " ")}</Td>
              <Td style={{ fontSize: 12 }}>{documentCategoryLabel(d.category)}</Td>
              <Td><button onClick={() => openClientDocument(d)} style={{ ...iconBtnStyle, padding: 0, fontFamily: sans, fontSize: 13, color: C.accentDark, textDecoration: "underline" }}>{d.fileName}</button></Td>
              <Td mono style={{ color: C.inkMuted }}>{fmtBytes(d.fileSize)}</Td>
              <Td style={{ color: C.inkMuted }}>{empName(d.uploadedBy)}</Td>
              <Td style={{ color: C.inkMuted }}>{d.note || "—"}</Td>
              <Td right><button onClick={async () => { if (await deleteClientDocument(d)) refetchDocuments(); }} style={iconBtnStyle} title="Delete"><Trash2 size={13} color={C.inkMuted} /></button></Td>
            </tr>
          ))}
        </TableShell>
      )}
    </Panel>
  );
}

/* ---------------------------------------------------------------------- */
/* Client card — Agreement tab (draft engagement letter from client data)  */
/* ---------------------------------------------------------------------- */

function buildAgreementHtml({ client, aml, firm, employees, contacts }) {
  const cur = currentFeeOf(client);
  const rep = aml?.representativeName || client.representativeName || contacts.find((c) => c.isPrimary)?.name || "[ΟΝΟΜΑΤΕΠΩΝΥΜΟ]";
  const repId = aml?.representativeIdNumber || client.representativeIdNumber || "[ΑΔΤ]";
  const repAfm = aml?.representativeAfm || client.representativeAfm || "[ΑΦΜ]";
  const start = client.contractStartDate || earliestHistoryDate(client.feeHistory) || client.goLiveDate || null;
  const signDate = client.contractSignedOn || start;
  const services = (client.contractServicesText || CONTRACT_DEFAULT_SERVICES).split("\n").map((l) => l.trim()).filter(Boolean);
  const e = escapeHtml;
  let n = 0; const art = (title) => `<h2>Άρθρο ${++n} – ${e(title)}</h2>`;
  let sub = 1; // 2.1 is the main service; the blocks below continue 2.2, 2.3, …
  const subHead = (title) => `<h3>2.${++sub} ${title}</h3>`;
  const payroll = client.contractPayrollBaseFee ? `
    ${subHead("Μισθοδοσία &amp; Εργατικά Θέματα")}
    <ul><li>Ενεργοποίηση module μισθοδοσίας και διαχείριση των ${client.contractPayrollIncluded || 2} πρώτων υπαλλήλων: ${e(fmtEurGr(client.contractPayrollBaseFee))} μηνιαίως πλέον ΦΠΑ.</li>
    ${client.contractPayrollExtraFee ? `<li>Κάθε επιπλέον υπάλληλος: ${e(fmtEurGr(client.contractPayrollExtraFee))} ανά άτομο μηνιαίως πλέον ΦΠΑ.</li>` : ""}</ul>` : "";
  // Annual work and start-up fees come from the client card (Fees tab), when switched on there.
  const annual = client.annualFeeEnabled && client.annualFee ? `
    ${subHead("Ετήσιες Εργασίες")}
    <p>Για την κατάρτιση των ετήσιων οικονομικών καταστάσεων (ισολογισμού) και την υποβολή της δήλωσης φορολογίας εισοδήματος:</p>
    <p><strong>Αμοιβή:</strong> ${e(fmtEurGr(client.annualFee))} ετησίως πλέον ΦΠΑ 24%.</p>` : "";
  const formation = client.formationFeeEnabled && client.formationFee ? `
    ${subHead("Εργασίες Έναρξης &amp; Σύστασης Επιχείρησης")}
    <p>Για τις εργασίες έναρξης δραστηριότητας και σύστασης της επιχείρησης (ενδεικτικά: υπηρεσία μιας στάσης, μισθωτήριο έδρας και υποβολή στην ΑΑΔΕ, εγγραφή στον ΕΦΚΑ, τραπεζικός λογαριασμός, ΓΕΜΗ, ηλεκτρονική τιμολόγηση):</p>
    <p><strong>Αμοιβή:</strong> ${e(fmtEurGr(client.formationFee))} εφάπαξ πλέον ΦΠΑ 24%.</p>` : "";
  const adhoc = client.contractHourlyRate ? `
    ${subHead("Πρόσθετες Υπηρεσίες (ad-hoc)")}
    <p>Για τυχόν εργασίες πέραν των ανωτέρω (π.χ. τροποποιήσεις καταστατικού κ.λπ.), ισχύει η παρακάτω χρέωση:</p>
    <p><strong>Χρέωση:</strong> ${e(fmtEurGr(client.contractHourlyRate))} ανά ώρα (gross rate) πλέον ΦΠΑ.</p>` : "";
  const otherFees = (client.contractOtherFeesText || "").trim() ? `
    ${subHead("Λοιπές Χρεώσεις &amp; Διευκρινίσεις")}
    <p>${e(client.contractOtherFeesText.trim()).replace(/\n/g, "<br/>")}</p>` : "";
  const notice = client.contractNoticeMonths ?? 3;
  const noticeText = notice === 1 ? "μηνιαία" : notice === 2 ? "δίμηνη" : notice === 3 ? "τρίμηνη" : notice === 6 ? "εξάμηνη" : `${notice} μηνών`;
  return `
  <h1 style="text-align:center">ΣΥΜΒΑΣΗ ΠΑΡΟΧΗΣ ΛΟΓΙΣΤΙΚΩΝ &amp; ΦΟΡΟΛΟΓΙΚΩΝ ΥΠΗΡΕΣΙΩΝ</h1>
  <p class="muted" style="text-align:center">Έκδοση ${client.contractVersion} · ${client.contractStatus === "signed" ? `Υπογράφηκε ${e(client.contractSignedOn)}` : `Σχέδιο — παράχθηκε ${new Date().toLocaleDateString("el-GR")}`}</p>
  <p>Σήμερα, την ${e(fmtGreekDate(signDate))}, μεταξύ:</p>
  <p>Αφενός της <strong>${e(client.name)}</strong>, με έδρα ${e(client.address || "[ΔΙΕΥΘΥΝΣΗ ΕΔΡΑΣ]")}, ΑΦΜ ${e(client.afm || "[ΑΦΜ]")}, Δ.Ο.Υ. ${e(client.doy || "[ΔΟΥ]")}, εφεξής καλούμενης «Εντολέας», με νόμιμο εκπρόσωπο τον/την ${e(rep)} με ΑΔΤ ${e(repId)} και ΑΦΜ ${e(repAfm)}</p>
  <p>και</p>
  <p>Αφετέρου της λογιστικής εταιρείας <strong>${e(firm.name)}</strong>, με έδρα ${e(firm.address)}, ΑΦΜ ${e(firm.afm)}, Δ.Ο.Υ. ${e(firm.doy)}, εφεξής καλούμενης «Ανάδοχος», με νόμιμο εκπρόσωπο τον ${e(firm.representative)}${firm.representativeId ? ` με ΑΔΤ ${e(firm.representativeId)}` : ""}${firm.representativeAfm ? ` και ΑΦΜ ${e(firm.representativeAfm)}` : ""},</p>
  <p>συμφωνούνται και γίνονται αμοιβαία αποδεκτά τα ακόλουθα:</p>
  ${art("Αντικείμενο Σύμβασης")}
  <p>Ο Ανάδοχος αναλαμβάνει την παροχή λογιστικών, φορολογικών και οικονομικών υπηρεσιών προς τον Εντολέα, σύμφωνα με την ισχύουσα φορολογική νομοθεσία. Η συνεργασία έχει διαρκή (ongoing) χαρακτήρα, με έναρξη ισχύος την ${e(fmtGreekDate(start))}.</p>
  ${art("Παρεχόμενες Υπηρεσίες")}
  <h3>2.1 Λογιστική – Φορολογική – Οικονομική Υποστήριξη</h3>
  <ul>${services.map((sv) => `<li>${e(sv)}</li>`).join("")}</ul>
  <p><strong>Αμοιβή:</strong> ${cur.fixedFee ? `${e(fmtEurGr(cur.fixedFee))} μηνιαίως πλέον ΦΠΑ 24%.` : "[ΠΟΣΟ] μηνιαίως πλέον ΦΠΑ 24%."}</p>
  ${payroll}${annual}${formation}${adhoc}${otherFees}
  <p>Σε περίπτωση μεταβολής των παρεχόμενων υπηρεσιών η σύμβαση θα τροποποιηθεί αναλόγως.</p>
  ${art("Οικονομικοί Όροι")}
  <p>${e(client.contractPaymentTerms || CONTRACT_DEFAULT_PAYMENT)}</p>
  ${art("Υποχρεώσεις Εντολέα")}
  <p>Ο Εντολέας υποχρεούται:</p>
  <ul><li>Να παρέχει έγκαιρα και με ακρίβεια όλα τα απαραίτητα στοιχεία και παραστατικά.</li><li>Να ενημερώνει τον Ανάδοχο για οποιαδήποτε μεταβολή αφορά τη δραστηριότητά του.</li><li>Να καταβάλλει εμπρόθεσμα τις συμφωνημένες αμοιβές.</li></ul>
  ${art("Προστασία Προσωπικών Δεδομένων (GDPR)")}
  <p>Τα συμβαλλόμενα μέρη δεσμεύονται να τηρούν τις διατάξεις του Γενικού Κανονισμού Προστασίας Δεδομένων (ΕΕ 2016/679 – GDPR) και της εθνικής νομοθεσίας.</p>
  <p>Ο Ανάδοχος επεξεργάζεται προσωπικά δεδομένα αποκλειστικά και μόνο για τους σκοπούς της παρούσας σύμβασης, λαμβάνοντας όλα τα απαραίτητα τεχνικά και οργανωτικά μέτρα για την ασφάλειά τους.</p>
  ${art("Μεταφορά Πληροφοριών")}
  <p>Όλα τα έγγραφα και οι πληροφορίες που αφορούν τη συμβατική μας σχέση και/ή τη διαχείριση των υποθέσεων του Εντολέα (εφεξής «Πληροφορίες») διαβιβάζονται με ασφαλή τρόπο μέσω ηλεκτρονικού ταχυδρομείου, συστημένης αλληλογραφίας/ταχυμεταφοράς ή μέσω ασφαλών διαδικτυακών εργαλείων διαμοιρασμού αρχείων εγκεκριμένων από τον Ανάδοχο (ενδεικτικά OneDrive, iManage κ.λπ.). Επιπλέον, εφαρμόζονται τα ακόλουθα μέτρα:</p>
  <p>(i) Οι πάροχοι ταχυδρομικών ή/και ταχυμεταφορικών υπηρεσιών διαθέτουν πιστοποιήσεις ποιότητας και ασφάλειας.<br/>(ii) Τηρείται αλυσίδα επιμέλειας (chain of custody) των Πληροφοριών σε όλα τα στάδια διαβίβασης.<br/>(iii) Πρόσβαση στις Πληροφορίες έχουν μόνο τα απολύτως απαραίτητα στελέχη του Αναδόχου.<br/>(iv) Εφαρμόζονται διαδικασίες ταυτοποίησης για ευαίσθητες ή/και εμπιστευτικές Πληροφορίες, εφόσον απαιτείται.<br/>(v) Χρησιμοποιείται κρυπτογράφηση ή/και προστασία με κωδικό πρόσβασης, εφόσον ζητηθεί.<br/>(vi) Όλα τα διαδικτυακά εργαλεία διαμοιρασμού αρχείων είναι εγκεκριμένα από το Τμήμα Ασφάλειας Πληροφοριών του Αναδόχου.</p>
  ${art("Εμπιστευτικότητα & Οδηγίες")}
  <p>Ο Εντολέας παρέχει στον Ανάδοχο τις απαραίτητες πληροφορίες για την εκτέλεση των υπηρεσιών. Ο Ανάδοχος βασίζεται στις πληροφορίες αυτές και σε δημόσια διαθέσιμες πηγές και δεν φέρει ευθύνη για την ακρίβεια ή πληρότητά τους.</p>
  <p>Ο Εντολέας οφείλει να γνωστοποιήσει εγγράφως τα πρόσωπα που είναι εξουσιοδοτημένα να παρέχουν οδηγίες ή να λαμβάνουν συμβουλές. Επισημαίνεται ότι, βάσει της ισχύουσας νομοθεσίας, στα τιμολόγια του Αναδόχου πρέπει να αναγράφεται ακριβής περιγραφή των παρεχόμενων υπηρεσιών.</p>
  <p>Ο Ανάδοχος τηρεί την εμπιστευτικότητα των Πληροφοριών σύμφωνα με την ελληνική νομοθεσία (ενδεικτικά: Ν.Δ. 3329/1955, Π.Δ. 226/1992, Κώδικας Δεοντολογίας ΣΟΕΛ, Ν. 3693/2008).</p>
  ${art("Χρήση Τεχνητής Νοημοσύνης")}
  <p>Ο Ανάδοχος δύναται να χρησιμοποιεί, κατά περίπτωση, συστήματα τεχνητής νοημοσύνης (AI), γενετικής τεχνητής νοημοσύνης ή άλλα αυτοματοποιημένα εργαλεία, αποκλειστικά για την υποστήριξη της παροχής των υπηρεσιών.</p>
  <p>Η χρήση αυτή γίνεται σύμφωνα με την ισχύουσα νομοθεσία, τους κανόνες δεοντολογίας, το GDPR και τις υποχρεώσεις επαγγελματικής ευθύνης. Ο Ανάδοχος παραμένει αποκλειστικά υπεύθυνος για την ποιότητα, την ακρίβεια και την τελική παράδοση των υπηρεσιών, οι οποίες υπόκεινται πάντοτε σε ανθρώπινο έλεγχο και επαγγελματική κρίση.</p>
  ${art("Νομικό Πλαίσιο")}
  <p>Ο Ανάδοχος είναι ${e(firm.legalFormText || "Ετερόρρυθμη Εταιρεία")} που λειτουργεί σύμφωνα με το ελληνικό δίκαιο. Η ευθύνη του Αναδόχου και των στελεχών του περιορίζεται σε περιπτώσεις βαριάς αμέλειας ή δόλου.</p>
  <p>Αποκλειστικά αρμόδια για την επίλυση κάθε διαφοράς είναι τα Δικαστήρια Αθηνών και εφαρμοστέο δίκαιο είναι το ελληνικό.</p>
  ${art("Καταπολέμηση Δωροδοκίας & Διαφθοράς")}
  <p>Ο Ανάδοχος εφαρμόζει πολιτική μηδενικής ανοχής σε πράξεις δωροδοκίας ή διαφθοράς και συμμορφώνεται πλήρως με την ισχύουσα νομοθεσία. Αντίστοιχη συμμόρφωση αναμένεται και από τον Εντολέα.</p>
  ${client.contractIncludeAml ? `${art("Δέουσα Επιμέλεια (ν. 4557/2018)")}
  <p>Ο Ανάδοχος, ως υπόχρεο πρόσωπο κατά τον ν. 4557/2018, εφαρμόζει μέτρα δέουσας επιμέλειας ως προς τον Εντολέα: ταυτοποίηση και επαλήθευση του Εντολέα, του νομίμου εκπροσώπου και των πραγματικών δικαιούχων, καταγραφή του σκοπού και της φύσης της συνεργασίας και συνεχή παρακολούθηση. Ο Εντολέας παρέχει τα έγγραφα της ΠΟΛ.1200/2018 και το αποδεικτικό καταχώρισης στο Κεντρικό Μητρώο Πραγματικών Δικαιούχων και γνωστοποιεί κάθε μεταβολή των στοιχείων του ή των πραγματικών δικαιούχων εντός εξήντα (60) ημερών. Ο Εντολέας ενημερώθηκε, κατά το άρθρο 31 παρ. 2 του ν. 4557/2018, για την επεξεργασία των δεδομένων του για τους σκοπούς αυτούς και για την τήρησή τους επί πέντε (5) έτη μετά τη λήξη της συνεργασίας. Σε αδυναμία ολοκλήρωσης της ταυτοποίησης, ο Ανάδοχος δύναται να μη συνάψει ή να διακόψει τη συνεργασία (άρθρο 13 παρ. 2).</p>` : ""}
  ${client.contractSpecialTerms ? `${art("Ειδικοί Όροι")}<p>${e(client.contractSpecialTerms).replace(/\n/g, "<br/>")}</p>` : ""}
  ${art("Λύση της Σύμβασης")}
  <p>Ο Εντολέας δύναται να καταγγείλει τη σύμβαση οποτεδήποτε, με έγγραφη καταγγελία και με ${e(noticeText)} προειδοποίηση.</p>
  <p>Ο Ανάδοχος δύναται να καταγγείλει τη σύμβαση σύμφωνα με τους κανόνες επαγγελματικής δεοντολογίας, αντιστοίχως με έγγραφη καταγγελία και με ${e(noticeText)} προειδοποίηση.</p>
  <p>Η λύση της σύμβασης δεν απαλλάσσει τον Εντολέα από την υποχρέωση εξόφλησης δεδουλευμένων αμοιβών και εξόδων. Η παράδοση των αρχείων πραγματοποιείται μετά την τακτοποίηση των σχετικών οφειλών.</p>
  <p style="margin-top:28px">Οι συμβαλλόμενοι,</p>
  <div class="sig"><div><strong>Ο Εντολέας</strong><br/>${e(rep.toUpperCase())}</div><div><strong>Ο Ανάδοχος</strong><br/>${e((firm.representative || firm.name).toUpperCase())}</div></div>
  <div class="footer">Παράχθηκε από το GLORIA στις ${new Date().toLocaleString("el-GR")} · ${client.contractStatus === "signed" ? "Τελικό κείμενο" : "Σχέδιο — δεν έχει υπογραφεί"}</div>`;
}

function ClientAgreementTab({ user, client, aml, firm, saveFirm, employees, contacts, ubos, documents, refetchClients, refetchDocuments }) {
  const fromClient = () => ({
    special: client.contractSpecialTerms || "", payment: client.contractPaymentTerms || "", servicesText: client.contractServicesText || CONTRACT_DEFAULT_SERVICES,
    payrollBase: client.contractPayrollBaseFee ?? "", payrollIncluded: client.contractPayrollIncluded ?? "", payrollExtra: client.contractPayrollExtraFee ?? "",
    hourlyRate: client.contractHourlyRate ?? "", otherFees: client.contractOtherFeesText || "", noticeMonths: client.contractNoticeMonths ?? 3, includeAml: client.contractIncludeAml !== false, startDate: client.contractStartDate || earliestHistoryDate(client.feeHistory) || "",
  });
  const [t, setT] = useState(fromClient);
  const [firmEdit, setFirmEdit] = useState(null);
  useEffect(() => { setT(fromClient()); }, [client.id, client.contractStatus, client.contractVersion]); // eslint-disable-line react-hooks/exhaustive-deps
  const dirty = JSON.stringify(t) !== JSON.stringify(fromClient());
  const special = t.special, payment = t.payment;
  const num = (v) => v === "" || v === null || v === undefined ? null : Number(v);
  const termsPatch = () => ({
    contract_special_terms: t.special.trim() || null, contract_payment_terms: t.payment.trim() || null, contract_services_text: t.servicesText.trim() === CONTRACT_DEFAULT_SERVICES.trim() ? null : (t.servicesText.trim() || null),
    contract_payroll_base_fee: num(t.payrollBase), contract_payroll_included: num(t.payrollIncluded), contract_payroll_extra_fee: num(t.payrollExtra), contract_hourly_rate: num(t.hourlyRate), contract_other_fees_text: t.otherFees.trim() || null,
    contract_notice_months: num(t.noticeMonths), contract_include_aml: !!t.includeAml, contract_start_date: t.startDate || null,
  });
  const previewClient = { ...client, contractSpecialTerms: t.special, contractPaymentTerms: t.payment, contractServicesText: t.servicesText, contractPayrollBaseFee: num(t.payrollBase), contractPayrollIncluded: num(t.payrollIncluded), contractPayrollExtraFee: num(t.payrollExtra), contractHourlyRate: num(t.hourlyRate), contractOtherFeesText: t.otherFees, contractNoticeMonths: num(t.noticeMonths), contractIncludeAml: !!t.includeAml, contractStartDate: t.startDate || null };
  const signedDocs = documents.filter((d) => d.category === "contract");
  const cur = currentFeeOf(client);
  const missing = [];
  if (!client.afm) missing.push("client ΑΦΜ");
  if (!client.address) missing.push("client address");
  if (!aml?.representativeName && !client.representativeName) missing.push("legal representative (Details tab)");
  if (!aml?.representativeIdNumber && !client.representativeIdNumber) missing.push("representative ΑΔΤ (Details tab)");
  if (!cur.fixedFee) missing.push("fixed fee (Fees tab)");

  async function saveTerms(extra = {}) {
    const { error } = await supabase.from("clients").update({ ...termsPatch(), ...extra }).eq("id", client.id);
    if (error) { window.alert(`Couldn't save: ${error.message}`); return; }
    refetchClients();
  }
  async function setStatus(status) {
    const patch = { contract_status: status };
    if (status === "signed") { const d = window.prompt("Signing date (YYYY-MM-DD)", client.contractSignedOn || toKey(TODAY)); if (d === null) return; patch.contract_signed_on = d || toKey(TODAY); }
    if (status === "draft" && client.contractStatus === "signed") patch.contract_version = client.contractVersion + 1;
    await saveTerms(patch);
    if (status === "signed") await applyOnboardingDeadlines(client.id, patch.contract_signed_on);
  }
  const preview = () => printHtml(`Σύμβαση — ${client.name}`, buildAgreementHtml({ client: previewClient, aml, firm: { ...DEFAULT_FIRM, ...firm }, employees, contacts }), { draft: false });

  return (
    <div>
      <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
        <StatCard label="Agreement" value={CONTRACT_STATUS_LABELS[client.contractStatus]} accent={client.contractStatus === "signed" ? C.accent : client.contractStatus === "none" ? undefined : C.warn} sub={client.contractSignedOn ? `Signed ${client.contractSignedOn} · v${client.contractVersion}` : `Version ${client.contractVersion}`} />
        <StatCard label="Monthly fee in the draft" value={cur.fixedFee ? fmtEur(cur.fixedFee) : "—"} sub="From the Fees tab" />
        <StatCard label="Signed copy on file" value={signedDocs.length ? "Yes" : "No"} accent={signedDocs.length ? C.accent : C.warn} sub={signedDocs.length ? signedDocs[0].fileName : "Upload under Documents → Signed agreement"} />
      </div>
      <Panel>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
          <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted, maxWidth: 640 }}>
            The draft is generated from the client card (details, representative, services from due diligence, fee from Fees). It carries a DRAFT watermark until you mark it signed.
            {missing.length > 0 && <div style={{ color: C.warn, marginTop: 4 }}>Missing for a complete draft: {missing.join(", ")}.</div>}
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Btn variant="secondary" onClick={() => setFirmEdit({ ...DEFAULT_FIRM, ...firm })}>Firm details</Btn>
            <Btn icon={Receipt} onClick={preview}>Preview / Save as PDF</Btn>
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
          <Field label="Start of services (Άρθρο 1)"><input type="date" style={inputStyle} value={t.startDate} onChange={(e) => setT({ ...t, startDate: e.target.value })} /></Field>
          <Field label="Notice period, months (Άρθρο 11)"><input type="number" min="0" style={inputStyle} value={t.noticeMonths} onChange={(e) => setT({ ...t, noticeMonths: e.target.value })} /></Field>
          <Field label="Monthly fee (from Fees tab)"><input style={{ ...inputStyle, background: C.bg }} value={cur.fixedFee ? fmtEurGr(cur.fixedFee) + " πλέον ΦΠΑ" : "—"} readOnly /></Field>
          <div style={{ gridColumn: "1 / span 3" }}><Field label="2.1 Services — one per line"><textarea style={{ ...textareaStyle, minHeight: 120 }} value={t.servicesText} onChange={(e) => setT({ ...t, servicesText: e.target.value })} /></Field></div>
          <Field label="2.2 Payroll module — monthly fee (€, empty = no payroll article)"><input type="number" min="0" style={inputStyle} value={t.payrollBase} onChange={(e) => setT({ ...t, payrollBase: e.target.value })} placeholder="e.g. 50" /></Field>
          <Field label="2.2 Employees included"><input type="number" min="0" style={inputStyle} value={t.payrollIncluded} onChange={(e) => setT({ ...t, payrollIncluded: e.target.value })} placeholder="e.g. 2" /></Field>
          <Field label="2.2 Each extra employee (€/month)"><input type="number" min="0" style={inputStyle} value={t.payrollExtra} onChange={(e) => setT({ ...t, payrollExtra: e.target.value })} placeholder="e.g. 15" /></Field>
          <Field label="2.3 Ad-hoc hourly rate (€, empty = no ad-hoc article)"><input type="number" min="0" style={inputStyle} value={t.hourlyRate} onChange={(e) => setT({ ...t, hourlyRate: e.target.value })} placeholder="e.g. 50" /></Field>
          <Field label="Annual fee (from Fees tab)"><input style={{ ...inputStyle, background: C.bg }} value={client.annualFeeEnabled && client.annualFee ? fmtEurGr(client.annualFee) + " ετησίως πλέον ΦΠΑ" : "Not applicable — switch on in Fees"} readOnly /></Field>
          <Field label="Start-up & formation fee (from Fees tab)"><input style={{ ...inputStyle, background: C.bg }} value={client.formationFeeEnabled && client.formationFee ? fmtEurGr(client.formationFee) + " εφάπαξ πλέον ΦΠΑ" : "Not applicable — switch on in Fees"} readOnly /></Field>
          <div />
          <div style={{ gridColumn: "1 / span 3" }}><Field label="Other charges & clarifications (free text — added as its own sub-article of Άρθρο 2)"><textarea style={{ ...textareaStyle, minHeight: 70 }} value={t.otherFees} onChange={(e) => setT({ ...t, otherFees: e.target.value })} placeholder="e.g. Χρέωση για υποβολή δήλωσης ΦΠΑ εκτός προθεσμίας, έξοδα ταχυμεταφοράς, εξαιρούμενες υπηρεσίες…" /></Field></div>
          <div style={{ gridColumn: "2 / span 2" }}><Field label="Άρθρο 3 — invoicing & payment terms (empty = template clause)"><input style={inputStyle} value={t.payment} onChange={(e) => setT({ ...t, payment: e.target.value })} placeholder={CONTRACT_DEFAULT_PAYMENT} /></Field></div>
          <div style={{ gridColumn: "1 / span 3" }}><Field label="Special terms (optional — added as a separate article)"><textarea style={{ ...textareaStyle, minHeight: 70 }} value={t.special} onChange={(e) => setT({ ...t, special: e.target.value })} /></Field></div>
          <label style={{ gridColumn: "1 / span 3", display: "flex", alignItems: "center", gap: 8, fontFamily: sans, fontSize: 12.5, color: C.ink }}><input type="checkbox" checked={!!t.includeAml} onChange={(e) => setT({ ...t, includeAml: e.target.checked })} /> Include the due-diligence article (ν. 4557/2018) — recommended by the compliance guide</label>
        </div>
        <div style={{ display: "flex", gap: 8, marginTop: 12, alignItems: "center", flexWrap: "wrap" }}>
          <Btn onClick={() => saveTerms()} disabled={!dirty}>Save terms</Btn>
          <span style={{ width: 1, height: 22, background: C.border, margin: "0 6px" }} />
          {client.contractStatus === "none" && <Btn variant="secondary" onClick={() => setStatus("draft")}>Mark as draft</Btn>}
          {(client.contractStatus === "draft") && <Btn variant="secondary" onClick={() => setStatus("sent")}>Mark as sent for signature</Btn>}
          {(client.contractStatus === "draft" || client.contractStatus === "sent") && <Btn variant="secondary" icon={Check} onClick={() => setStatus("signed")}>Mark as signed</Btn>}
          {client.contractStatus === "signed" && <Btn variant="ghost" onClick={() => setStatus("draft")}>New version (back to draft)</Btn>}
          {dirty && <span style={{ fontFamily: sans, fontSize: 12, color: C.warn }}>Unsaved changes</span>}
        </div>
      </Panel>
      {firmEdit && (
        <Modal title="Firm details (used in every agreement)" onClose={() => setFirmEdit(null)} width={480}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            {[["name", "Firm name"], ["legalFormText", "Legal form (as written in Άρθρο 9)"], ["afm", "ΑΦΜ"], ["doy", "ΔΟΥ"], ["gemi", "ΓΕΜΗ"], ["representative", "Legal representative"], ["representativeId", "Representative ΑΔΤ"], ["representativeAfm", "Representative ΑΦΜ"], ["email", "Email"], ["phone", "Phone"]].map(([k, l]) => (
              <Field key={k} label={l}><input style={inputStyle} value={firmEdit[k] || ""} onChange={(e) => setFirmEdit({ ...firmEdit, [k]: e.target.value })} /></Field>
            ))}
            <div style={{ gridColumn: "1 / span 2" }}><Field label="Address (seat)"><input style={inputStyle} value={firmEdit.address || ""} onChange={(e) => setFirmEdit({ ...firmEdit, address: e.target.value })} /></Field></div>
          </div>
          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 12 }}><Btn variant="secondary" onClick={() => setFirmEdit(null)}>Cancel</Btn><Btn onClick={async () => { await saveFirm(firmEdit); setFirmEdit(null); }}>Save</Btn></div>
        </Modal>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Due diligence file as printable PDF (time-stamped)                      */
/* ---------------------------------------------------------------------- */

// Pre-filled KYC questionnaire (ν. 4557/2018, ΠΟΛ.1200/2018) — printed from the Due diligence tab, signed by the client, then uploaded back.
function buildKycQuestionnaireHtml({ client, aml, ubos, firm = DEFAULT_FIRM }) {
  const a = aml || {};
  const e = escapeHtml;
  const type = a.entityType || amlEntityTypeFor(client.legalForm);
  const legal = type !== "natural";
  const v = (x, w = 28) => (x ? `<strong>${e(x)}</strong>` : "&nbsp;".repeat(w));
  const box = (on) => `<span class="cb" contenteditable="false" style="cursor:pointer;user-select:none">${on ? "☒" : "☐"}</span>`;
  const docs = a.documents || {};
  const rf = new Set(a.riskFactors || []);
  const checked = !!a.sanctionsCheckedOn; // PEP / sanctions screening done → "no" answers can be ticked from the file
  const noUbo = !rf.has("a_ubo_unknown");
  const row = (l, val) => `<tr><td style="width:34%;background:#f6f6f6"><strong>${l}</strong></td><td>${val || "&nbsp;"}</td></tr>`;
  // v: true → ΝΑΙ ticked, false → ΟΧΙ ticked, undefined → both blank (client answers)
  const yn = (l, v) => `<tr><td>${l}</td><td style="width:60px;text-align:center">${box(v === true)}</td><td style="width:60px;text-align:center">${box(v === false)}</td></tr>`;
  const ynHead = (h) => `<tr><th>${h}</th><th style="width:60px;text-align:center">ΝΑΙ</th><th style="width:60px;text-align:center">ΟΧΙ</th></tr>`;
  const myUbos = (ubos || []).filter((u) => u.clientId === client.id);
  const uboRows = (myUbos.length ? myUbos : [null, null, null]).map((u) => `<tr><td>${u ? e(u.name) : "&nbsp;"}</td><td>${u ? e([u.afm, u.idNumber].filter(Boolean).join(" / ")) : ""}</td><td>${u ? e([u.nationality, u.residenceCountry].filter(Boolean).join(" · ")) : ""}</td><td>${u && u.ownershipPct != null ? e(u.ownershipPct + "%") : ""}</td><td>${u ? e(u.controlType) : ""}</td></tr>`).join("");
  const sig = `<table style="margin-top:26px"><tr><th style="width:50%">Για τον πελάτη${legal ? " (νόμιμος εκπρόσωπος)" : ""}</th><th>Για το γραφείο (${e(firm.name)})</th></tr><tr><td style="height:90px;vertical-align:top">Ονοματεπώνυμο:<br/>Ιδιότητα:<br/>Ημερομηνία:<br/><br/>Υπογραφή${legal ? " και σφραγίδα" : ""}:</td><td style="vertical-align:top">Ονοματεπώνυμο:<br/>Ιδιότητα:<br/>Ημερομηνία:<br/><br/>Υπογραφή:</td></tr></table>`;
  const pepRows = ["Αρχηγός κράτους, αρχηγός κυβέρνησης, υπουργός, αναπληρωτής υπουργού ή υφυπουργός", "Μέλος κοινοβουλίου ή αντίστοιχου νομοθετικού σώματος", "Μέλος διοικητικού οργάνου πολιτικού κόμματος", "Μέλος ανώτατου δικαστηρίου, συνταγματικού δικαστηρίου ή άλλου υψηλού επιπέδου δικαστικού οργάνου", "Μέλος ελεγκτικού δικαστηρίου (περιλαμβάνεται το Ελεγκτικό Συνέδριο) ή διοικητικού συμβουλίου κεντρικής τράπεζας", "Πρέσβης, επιτετραμμένος ή ανώτερος αξιωματικός των ενόπλων δυνάμεων", "Μέλος διοικητικού, διαχειριστικού ή εποπτικού οργάνου κρατικής επιχείρησης", "Διευθυντής, αναπληρωτής διευθυντής ή μέλος ΔΣ διεθνούς οργανισμού", "Μέλος της οικογένειας (σύζυγος, τέκνα και σύζυγοί τους, γονείς) ή στενός συνεργάτης των ανωτέρω"].map((l) => yn(l, checked && (a.pepStatus || "none") === "none" ? false : undefined)).join("");
  return `
  <h1 style="color:#1f5e3b">ΕΡΩΤΗΜΑΤΟΛΟΓΙΟ ΓΝΩΡΙΜΙΑΣ ΠΕΛΑΤΗ (KYC) ${legal ? "ΓΙΑ ΝΟΜΙΚΟ ΠΡΟΣΩΠΟ" : "ΓΙΑ ΦΥΣΙΚΟ ΠΡΟΣΩΠΟ"}</h1>
  <p class="muted">ν. 4557/2018 · ταυτοποίηση κατά την ΠΟΛ.1200/2018, Πίνακες ${legal ? "Β, Γ, Δ και Πίνακας Α για τον εκπρόσωπο" : "Α"}. Συμπληρώνεται πριν από τη σύναψη της επιχειρηματικής σχέσης.</p>
  <table>${row("Πελάτης", v(client.name))}${row("Ημερομηνία συμπλήρωσης", "&nbsp;")}${row("Υπεύθυνος συμμόρφωσης", "&nbsp;")}</table>

  <h2>Α. ΣΤΟΙΧΕΙΑ ΤΑΥΤΟΠΟΙΗΣΗΣ ${legal ? "ΤΟΥ ΝΟΜΙΚΟΥ ΠΡΟΣΩΠΟΥ" : "ΤΟΥ ΠΕΛΑΤΗ"}</h2>
  ${legal ? `<table>${ynHead("Είδος προσώπου")}<tr><td>Κεφαλαιουχική εταιρεία (Α.Ε., Ε.Π.Ε., Ι.Κ.Ε., μονοπρόσωπη)</td><td style="text-align:center">${box(type === "capital")}</td><td style="text-align:center">${box(type !== "capital")}</td></tr><tr><td>Προσωπική εταιρεία (Ο.Ε., Ε.Ε.)</td><td style="text-align:center">${box(type === "partnership")}</td><td style="text-align:center">${box(type !== "partnership")}</td></tr><tr><td>Λοιπό νομικό πρόσωπο ή οντότητα</td><td style="text-align:center">${box(type === "other")}</td><td style="text-align:center">${box(type !== "other")}</td></tr></table>` : ""}
  <table class="kv">
    ${row(legal ? "Επωνυμία και διακριτικός τίτλος" : "Ονοματεπώνυμο", v(client.name))}
    ${row("Νομική μορφή", v(client.legalForm))}
    ${row(legal ? "Έδρα (διεύθυνση)" : "Διεύθυνση κατοικίας / επαγγελματικής εγκατάστασης", v(client.address))}
    ${row("Α.Φ.Μ. και Δ.Ο.Υ.", v([client.afm, client.doy].filter(Boolean).join(" / ")))}
    ${legal ? row("Αριθμός Γ.Ε.ΜΗ. · ημερομηνία σύστασης", v([a.gemiNumber, a.incorporationDate].filter(Boolean).join(" · "))) : ""}
    ${row("Κύρια δραστηριότητα / Κ.Α.Δ.", v(a.activity))}
    ${row("Τηλέφωνο και email επικοινωνίας", v([client.phone, client.email].filter(Boolean).join(" · ")))}
  </table>
  <p class="muted">Έγγραφα που προσκομίζονται: ${legal ? "καταστατικό και τροποποιήσεις, πιστοποιητικό Γ.Ε.ΜΗ./ανακοίνωση, πράξη διορισμού διαχειριστών, βεβαίωση έναρξης εργασιών, δήλωση πελάτη." : "ταυτότητα ή διαβατήριο, βεβαίωση έναρξης/μεταβολής εργασιών, απόδειξη διεύθυνσης (λογαριασμός κοινής ωφελείας ή μίσθωση)."}</p>

  <h2>${legal ? "Β. ΝΟΜΙΜΟΣ ΕΚΠΡΟΣΩΠΟΣ (ΠΙΝΑΚΑΣ Α)" : "Β. ΣΤΟΙΧΕΙΑ ΤΑΥΤΟΤΗΤΑΣ (ΠΙΝΑΚΑΣ Α)"}</h2>
  <table class="kv">
    ${row("Ονοματεπώνυμο και πατρώνυμο", v(a.representativeName))}
    ${row("Αριθμός δελτίου ταυτότητας / διαβατηρίου και εκδούσα αρχή", v([a.representativeIdType, a.representativeIdNumber].filter(Boolean).join(" · ")))}
    ${row("Α.Φ.Μ. εκπροσώπου", v(a.representativeAfm))}
    ${row("Ημερομηνία και τόπος γέννησης · ιθαγένεια", "&nbsp;")}
    ${legal ? row("Ιδιότητα (διαχειριστής, πρόεδρος Δ.Σ., εταίρος κ.λπ.)", "&nbsp;") : ""}
  </table>

  ${legal ? `<h2>Γ. ΠΡΑΓΜΑΤΙΚΟΙ ΔΙΚΑΙΟΥΧΟΙ</h2>
  <table><tr><th>Ονοματεπώνυμο</th><th>ΑΦΜ / Ταυτότητα</th><th>Ιθαγένεια · Κατοικία</th><th>Ποσοστό</th><th>Τρόπος ελέγχου</th></tr>${uboRows}</table>
  <table>${ynHead("Έλεγχοι")}
    ${yn("Η εταιρεία έχει καταχωρίσει τους πραγματικούς δικαιούχους στο Κεντρικό Μητρώο Πραγματικών Δικαιούχων (ΚΜΠΔ)", docs.ubo_registry?.collected ? true : undefined)}
    ${yn("Προσκομίστηκε μητρώο μετόχων / εταίρων", docs.shareholders?.collected ? true : undefined)}
    ${yn("Υπάρχει σύνθετη δομή ιδιοκτησίας ή εικονικοί μέτοχοι (nominees)", rf.has("a_nominee") || rf.has("a_complex") ? true : docs.shareholders?.collected ? false : undefined)}
    ${yn("Δεν εντοπίστηκε φυσικό πρόσωπο που κατέχει πάνω από 25% και ορίστηκε ο ανώτερος διευθυντικός παράγοντας", rf.has("a_ubo_unknown") ? true : myUbos.length ? false : undefined)}
  </table>
  <p class="muted">Ο πελάτης δεσμεύεται να ενημερώνει το γραφείο για κάθε μεταβολή των πραγματικών δικαιούχων και να την καταχωρίζει στο ΚΜΠΔ εντός εξήντα (60) ημερών.</p>` : ""}

  <h2>${legal ? "Δ" : "Γ"}. ΕΞΟΥΣΙΟΔΟΤΗΜΕΝΟ ΠΡΟΣΩΠΟ</h2>
  <table class="kv">${row("Ονοματεπώνυμο εξουσιοδοτημένου", v(a.proxyName))}${row("Α.Φ.Μ. / ταυτότητα", v(a.proxyIdNumber))}${row("Έγγραφο νομιμοποίησης (πληρεξούσιο, εξουσιοδότηση)", "&nbsp;")}</table>

  <h2>${legal ? "Ε" : "Δ"}. ΠΟΛΙΤΙΚΑ ΕΚΤΕΘΕΙΜΕΝΑ ΠΡΟΣΩΠΑ</h2>
  <p class="muted">Άρθρο 3 παρ. 9–11 ν. 4557/2018. Ισχύει για τον ${legal ? "εκπρόσωπο και τους πραγματικούς δικαιούχους" : "πελάτη"}, καθώς και για την οικογένεια και τους στενούς συνεργάτες τους.</p>
  <table>${ynHead("Κατηγορία")}${pepRows}</table>
  ${(a.pepStatus || "none") !== "none" ? `<p><strong>Σημείωση γραφείου:</strong> στον φάκελο έχει καταχωριστεί ΠΕΠ (${e(a.pepStatus)}${a.pepDetails ? " · " + e(a.pepDetails) : ""}). Σημειώστε την αντίστοιχη κατηγορία.</p>` : ""}

  <h2>${legal ? "ΣΤ" : "Ε"}. ΓΕΩΓΡΑΦΙΚΟΙ ΠΑΡΑΓΟΝΤΕΣ</h2>
  <table>${ynHead("Ερώτηση")}
    ${yn("Δραστηριότητα, κατοικία ή περιουσιακά στοιχεία σε τρίτη χώρα υψηλού κινδύνου (κατάλογος ΕΕ / FATF)", rf.has("b_eu_list") || rf.has("b_weak_aml") ? true : checked ? false : undefined)}
    ${yn("Σχέση με χώρα που υπόκειται σε κυρώσεις ή εμπάργκο", rf.has("b_sanctions") || rf.has("b_terror") ? true : checked ? false : undefined)}
    ${yn("Οι συναλλαγές εκτελούνται με μεταφορά κεφαλαίων από/προς το εξωτερικό")}
  </table>
  <table class="kv">${row("Χώρες δραστηριότητας / συναλλαγών", v(a.countries))}</table>

  <h2>${legal ? "Ζ" : "ΣΤ"}. ΣΚΟΠΟΣ ΚΑΙ ΦΥΣΗ ΤΗΣ ΣΥΝΑΛΛΑΓΙΚΗΣ ΣΧΕΣΗΣ</h2>
  <table class="kv">
    ${row("Υπηρεσίες που ζητούνται", v((a.services || []).join(", ")))}
    ${row("Σκοπός και φύση της σχέσης", v(a.relationshipPurpose))}
    ${row("Αναμενόμενος όγκος και είδος συναλλαγών", "&nbsp;")}
    ${row("Προέλευση κεφαλαίων και περιουσίας", v(a.sourceOfFunds))}
    ${row("Τρόπος πληρωμής αμοιβών (τραπεζικό έμβασμα, κάρτα, μετρητά)", "&nbsp;")}
  </table>

  <h2>${legal ? "Η" : "Ζ"}. ΕΝΗΜΕΡΩΣΗ ΓΙΑ ΤΗΝ ΕΠΕΞΕΡΓΑΣΙΑ ΠΡΟΣΩΠΙΚΩΝ ΔΕΔΟΜΕΝΩΝ (άρθρο 31 παρ. 2 ν. 4557/2018)</h2>
  <p>Η ${e(firm.name)}, ως υπόχρεο πρόσωπο, επεξεργάζεται τα προσωπικά δεδομένα που περιλαμβάνονται στο παρόν και στα συνημμένα έγγραφα αποκλειστικά για την πρόληψη της νομιμοποίησης εσόδων από εγκληματικές δραστηριότητες και της χρηματοδότησης της τρομοκρατίας. Τα δεδομένα δεν χρησιμοποιούνται για άλλους σκοπούς.</p>
  <p>Τα δεδομένα φυλάσσονται για πέντε (5) έτη μετά τη λήξη της επιχειρηματικής σχέσης, εκτός αν άλλη διάταξη επιβάλλει μεγαλύτερη διάρκεια. Μπορούν να διαβιβαστούν στην Αρχή Καταπολέμησης της Νομιμοποίησης Εσόδων από Εγκληματικές Δραστηριότητες και σε άλλες αρμόδιες αρχές, όταν ο νόμος το επιβάλλει. Έχετε δικαίωμα πρόσβασης και διόρθωσης, ενώ το δικαίωμα διαγραφής δεν ασκείται όσο διαρκεί η υποχρεωτική τήρηση. Μπορείτε να απευθυνθείτε στο γραφείο και να προσφύγετε στην Αρχή Προστασίας Δεδομένων Προσωπικού Χαρακτήρα.</p>
  ${legal ? `<p>Η ενημέρωση αυτή αφορά και τα φυσικά πρόσωπα που συνδέονται με την εταιρεία (πραγματικοί δικαιούχοι, εκπρόσωποι). Ο εκπρόσωπος δηλώνει ότι τα ενημέρωσε.</p>` : ""}

  <h2>ΔΗΛΩΣΕΙΣ ΚΑΙ ΥΠΟΓΡΑΦΕΣ</h2>
  <p>Ο υπογράφων δηλώνει ότι τα ανωτέρω στοιχεία είναι αληθή και ακριβή και αναλαμβάνει την υποχρέωση να γνωστοποιεί χωρίς καθυστέρηση κάθε μεταβολή τους${legal ? ", ιδίως ως προς τους πραγματικούς δικαιούχους, τη δομή διοίκησης, τη δραστηριότητα και την ιδιότητα πολιτικώς εκτεθειμένου προσώπου" : ", ιδίως ως προς τη διεύθυνση, τη δραστηριότητα και την ιδιότητα πολιτικώς εκτεθειμένου προσώπου"}. Δηλώνει επίσης ότι έλαβε την ανωτέρω ενημέρωση πριν από τη σύναψη της επιχειρηματικής σχέσης.</p>
  ${sig}
  <div class="footer">KYC ${legal ? "νομικού" : "φυσικού"} προσώπου · ν. 4557/2018 · ${e(client.name)}</div>`;
}

function buildDueDiligenceHtml({ client, aml, ubos, documents, employees, tasks }) {
  const empName = (id) => employees.find((e) => e.id === id)?.name || "—";
  const kv = (rows) => `<table class="kv">${rows.map(([k, v]) => `<tr><td>${escapeHtml(k)}</td><td>${escapeHtml(v || "—")}</td></tr>`).join("")}</table>`;
  const a = aml || {};
  const assessed = amlAssess(a.riskFactors || []);
  const level = a.riskLevel || assessed.level;
  const chosen = AML_RISK_FACTORS.filter((f) => (a.riskFactors || []).includes(f.key));
  const docsForType = AML_DOCUMENTS.filter((d) => d.types.includes(a.entityType || amlEntityTypeFor(client.legalForm)));
  const files = documents.filter((d) => d.category !== "other" && d.category !== "authorisation");
  const amlTasks = tasks.filter((t) => t.clientId === client.id && t.kind === "onboarding" && t.stage === "aml");
  const now = new Date();
  return `
  <h1>Φάκελος δέουσας επιμέλειας πελάτη — ${escapeHtml(client.name)}</h1>
  <p class="muted">ν. 4557/2018 · ΠΟΛ.1200/2018 · Εκτύπωση: ${now.toLocaleString("el-GR")} · Κατάσταση φακέλου: ${escapeHtml(AML_STATUS_LABELS[a.status || "not_started"])}${a.completedOn ? ` (${a.completedOn}, ${escapeHtml(empName(a.completedBy))})` : ""}</p>
  <h2>1. Ταυτοποίηση πελάτη</h2>
  ${kv([["Επωνυμία", client.name], ["Νομική μορφή · Πίνακας ΠΟΛ.1200", `${client.legalForm || "—"} · ${AML_ENTITY_TYPES[a.entityType || amlEntityTypeFor(client.legalForm)]}`], ["ΑΦΜ / ΔΟΥ", [client.afm, client.doy].filter(Boolean).join(" / ")], ["ΓΕΜΗ · Ίδρυση", [a.gemiNumber, a.incorporationDate].filter(Boolean).join(" · ")], ["Έδρα", client.address], ["Νόμιμος εκπρόσωπος", a.representativeName], ["ΑΦΜ εκπροσώπου", a.representativeAfm], ["Έγγραφο ταυτότητας", [a.representativeIdType, a.representativeIdNumber].filter(Boolean).join(" · ")], ["Εξουσιοδοτημένο πρόσωπο", a.proxyName ? `${a.proxyName} · ${a.proxyIdNumber || ""}` : "Δεν ενεργεί μέσω τρίτου"]])}
  <h2>2. Πραγματικοί δικαιούχοι</h2>
  ${ubos.length === 0 ? "<p>Δεν έχουν καταχωριστεί.</p>" : `<table><tr><th>Ονοματεπώνυμο</th><th>ΑΦΜ / Ταυτότητα</th><th>Εθνικότητα · Κατοικία</th><th>Ποσοστό / έλεγχος</th><th>ΠΕΠ</th><th>Επαλήθευση</th></tr>${ubos.map((u) => `<tr><td>${escapeHtml(u.name)}</td><td>${escapeHtml([u.afm, u.idNumber].filter(Boolean).join(" / "))}</td><td>${escapeHtml([u.nationality, u.residenceCountry].filter(Boolean).join(" · "))}</td><td>${u.ownershipPct !== null ? u.ownershipPct + "%" : ""} ${escapeHtml(u.controlType || "")}</td><td>${u.isPep ? "ΝΑΙ" : "Όχι"}</td><td>${escapeHtml([u.verifiedOn, u.verificationDoc].filter(Boolean).join(" · ") || "Μη επαληθευμένος")}</td></tr>`).join("")}</table>`}
  ${a.uboActionsRecord ? `<p><strong>Αρχείο ενεργειών (άρ. 3 παρ. 17 αβ΄):</strong> ${escapeHtml(a.uboActionsRecord)}</p>` : ""}
  <h2>3. Έλεγχος ΠΕΠ και κυρώσεων</h2>
  ${kv([["ΠΕΠ", a.pepStatus && a.pepStatus !== "none" ? `ΝΑΙ — ${a.pepStatus} · ${a.pepDetails || ""}` : "Όχι"], ["Έλεγχος λιστών", a.sanctionsCheckedOn], ["Αποτέλεσμα", a.sanctionsResult]])}
  <h2>4. Σκοπός και φύση της σχέσης</h2>
  ${kv([["Δραστηριότητα", a.activity], ["Χώρες", a.countries], ["Υπηρεσίες", (a.services || []).join(", ")], ["Σκοπός / αναμενόμενη συναλλακτική εικόνα", a.relationshipPurpose], ["Πηγές πληροφοριών", a.infoSources], ["Πηγή πλούτου / προέλευση κεφαλαίων", a.sourceOfFunds]])}
  <h2>5. Αξιολόγηση κινδύνου (Μοντέλο Α — 4 άξονες)</h2>
  <table><tr><th>Άξονας</th><th>Παράγοντας</th><th>Παράρτημα / διάταξη</th><th>Βαθμοί</th><th>Τεκμηρίωση</th></tr>${chosen.length === 0 ? "<tr><td colspan='5'>Κανένας παράγοντας</td></tr>" : chosen.map((f) => `<tr><td>${f.axis}</td><td>${escapeHtml(f.label)}</td><td>${escapeHtml(f.ref)}</td><td>${f.override ? "ΥΠΕΡΙΣΧΥΕΙ" : (f.points > 0 ? "+" : "") + f.points}</td><td>${escapeHtml((a.riskFactorNotes || {})[f.key] || "")}</td></tr>`).join("")}</table>
  ${kv([["Άθροισμα (Μοντέλο Α)", `${assessed.score >= 0 ? "+" : ""}${assessed.score}`], ["Βαθμός 0-100 (Μοντέλο Β)", String(amlAssessB(a.riskFactors || []).score)], ["Τύπος / έκδοση φύλλου", `${AML_ASSESSMENT_TYPES[a.assessmentType || "initial"]} · έκδοση ${a.sheetVersion || 1} · ${a.assessmentDate || ""}`], ["Κατάταξη", `${AML_LEVEL_LABELS[level]}${a.riskOverride && !assessed.forced ? ` (τεκμηριωμένη απόκλιση από ${AML_LEVEL_LABELS[assessed.level]}: ${a.riskOverrideReason || ""})` : ""}`], ["Επίπεδο δέουσας επιμέλειας", AML_DD_LABELS[a.ddLevel || AML_DD_FOR_LEVEL[level]]], ["Έγκριση ανώτερου στελέχους", a.approvedBy ? `${empName(a.approvedBy)} · ${a.approvedOn || ""}` : (level === "high" ? "ΕΚΚΡΕΜΕΙ" : "Δεν απαιτείται")]])}
  <h2>6. Έγγραφα φακέλου</h2>
  <table><tr><th>Έγγραφο</th><th>Συλλέχθηκε</th><th>Ημερομηνία</th><th>Αρχείο στο σύστημα</th></tr>${docsForType.map((d) => { const st = (a.documents || {})[d.key] || {}; const f = files.filter((x) => x.category === d.key); return `<tr><td>${escapeHtml(d.label)}</td><td>${st.collected ? "ΝΑΙ" : "—"}</td><td>${escapeHtml(st.date || "")}</td><td>${f.map((x) => `${escapeHtml(x.fileName)} (${x.uploadedAt.slice(0, 10)})`).join("<br/>") || "—"}</td></tr>`; }).join("")}</table>
  <h2>7. Βήματα διαδικασίας</h2>
  ${amlTasks.length === 0 ? "<p>Δεν δημιουργήθηκαν tasks δέουσας επιμέλειας.</p>" : `<table><tr><th>Βήμα</th><th>Ανατέθηκε</th><th>Προθεσμία</th><th>Κατάσταση</th><th>Ολοκληρώθηκε</th></tr>${amlTasks.map((t) => `<tr><td>${escapeHtml(t.title)}</td><td>${escapeHtml(empName(t.assignedTo))}</td><td>${escapeHtml(t.dueDate || "")}</td><td>${escapeHtml(TASK_STATUSES[t.status])}</td><td>${escapeHtml(t.doneAt || "")}</td></tr>`).join("")}</table>`}
  <h2>8. Απόφαση και επανεξέταση</h2>
  ${kv([["Απόφαση", a.decision === "accept" ? "Ανάληψη" : a.decision === "decline" ? `Άρνηση — ${a.decisionReason || ""}` : "Εκκρεμεί"], ["Ολοκλήρωση φακέλου", a.completedOn ? `${a.completedOn} · ${empName(a.completedBy)}` : "—"], ["Επόμενη επικαιροποίηση", a.nextReviewOn], ["Σημειώσεις", a.notes]])}
  <div class="sig"><div>Υπεύθυνος συμμόρφωσης<br/>${escapeHtml(empName(a.completedBy))}</div><div>Ημερομηνία<br/>${escapeHtml(a.completedOn || "")}</div></div>
  <div class="footer">Το παρόν αποτελεί εκτύπωση του ηλεκτρονικού φακέλου του πελάτη από το GLORIA την ${now.toLocaleString("el-GR")}. Τηρείται για 5 έτη μετά τη λήξη της επιχειρηματικής σχέσης (άρ. 30 ν. 4557/2018).</div>`;
}

/* ---------------------------------------------------------------------- */
/* Client card — Due diligence tab (KYC / AML, ν. 4557/2018)                */
/* Structure follows the Taxheaven guide's new-client flow (Β.8) and risk  */
/* model (Β.5.2), with documents per ΠΟΛ.1200/2018 tables A–D.             */
/* ---------------------------------------------------------------------- */

function DueDiligenceTab({ user, client, employees, aml, ubos, tasks, documents = [], refetchDocuments, refetchDueDiligence, refetchTasks }) {
  const defaultEntity = amlEntityTypeFor(client.legalForm);
  const blank = {
    clientId: client.id, clientType: defaultEntity === "natural" ? "natural" : "legal", entityType: defaultEntity, gemiNumber: "", incorporationDate: "",
    representativeName: client.representativeName || "", representativeIdType: client.representativeIdType || "", representativeIdNumber: client.representativeIdNumber || "", representativeAfm: client.representativeAfm || "", activity: "", countries: "Ελλάδα",
    relationshipPurpose: "", infoSources: "", services: [], sourceOfFunds: "", pepStatus: "none", pepDetails: "", sanctionsCheckedOn: "", sanctionsResult: "",
    proxyName: "", proxyIdNumber: "", riskFactors: ["b_eu"], riskScore: 0, riskLevel: "medium", riskOverride: "", riskOverrideReason: "", ddLevel: "standard", documents: {},
    status: "not_started", completedOn: "", completedBy: null, approvedBy: null, approvedOn: "", nextReviewOn: "", notes: "", decision: "", decisionReason: "", uboActionsRecord: "",
    riskFactorNotes: {}, assessmentType: "initial", assessmentDate: toKey(TODAY), sheetVersion: 1, riskScoreB: 0, uboChangeDate: "", uboFiledDate: "",
  };
  const withRep = (a) => ({ ...a, entityType: a.entityType || defaultEntity, representativeName: a.representativeName || client.representativeName || "", representativeAfm: a.representativeAfm || client.representativeAfm || "", representativeIdType: a.representativeIdType || client.representativeIdType || "", representativeIdNumber: a.representativeIdNumber || client.representativeIdNumber || "" });
  const init = aml ? withRep(aml) : blank;
  const [f, setF] = useState(init);
  useEffect(() => { setF(aml ? withRep(aml) : blank); }, [aml?.clientId, client.id]); // eslint-disable-line react-hooks/exhaustive-deps
  const [uboModal, setUboModal] = useState(null);
  const [u, setU] = useState({});

  const assessed = amlAssess(f.riskFactors);
  const assessedB = amlAssessB(f.riskFactors);
  const effectiveLevel = assessed.forced ? "high" : (f.riskOverride || assessed.level);
  const ddLevel = AML_DD_FOR_LEVEL[effectiveLevel];
  // "Πρακτικός κανόνας": an Annex II factor with a non-high classification
  // needs a written justification on the sheet.
  const annexIIPresent = AML_RISK_FACTORS.some((rf) => f.riskFactors.includes(rf.key) && (rf.points || 0) > 0);
  const needsJustification = annexIIPresent && effectiveLevel !== "high" && !(f.riskOverrideReason || "").trim();
  const isLegal = f.entityType !== "natural";
  const setFactorNote = (k, v) => set({ riskFactorNotes: { ...(f.riskFactorNotes || {}), [k]: v } });
  const set = (patch) => setF((prev) => ({ ...prev, ...patch }));
  const toggleFactor = (k) => set({ riskFactors: f.riskFactors.includes(k) ? f.riskFactors.filter((x) => x !== k) : [...f.riskFactors, k] });
  const toggleService = (k) => set({ services: f.services.includes(k) ? f.services.filter((x) => x !== k) : [...f.services, k] });
  const setDoc = (k, patch) => set({ documents: { ...f.documents, [k]: { ...(f.documents[k] || {}), ...patch } } });
  const docsForType = AML_DOCUMENTS.filter((d) => d.types.includes(f.entityType));
  const docsRequired = docsForType.filter((d) => !d.optional && (!d.enhancedOnly || ddLevel === "enhanced"));
  const docsCollected = docsRequired.filter((d) => f.documents[d.key]?.collected).length;
  const empName = (id) => employees.find((e) => e.id === id)?.name || "—";
  const dirty = JSON.stringify(f) !== JSON.stringify(init);

  async function save(extra = {}) {
    const row = {
      client_id: client.id, client_type: isLegal ? "legal" : "natural", entity_type: f.entityType, gemi_number: f.gemiNumber || null, incorporation_date: f.incorporationDate || null,
      representative_name: f.representativeName || null, representative_id_type: f.representativeIdType || null, representative_id_number: f.representativeIdNumber || null, representative_afm: f.representativeAfm || null,
      activity: f.activity || null, countries: f.countries || null, relationship_purpose: f.relationshipPurpose || null, info_sources: f.infoSources || null, services: f.services,
      source_of_funds: f.sourceOfFunds || null, pep_status: f.pepStatus, pep_details: f.pepDetails || null, sanctions_checked_on: f.sanctionsCheckedOn || null, sanctions_result: f.sanctionsResult || null,
      proxy_name: f.proxyName || null, proxy_id_number: f.proxyIdNumber || null,
      risk_factors: f.riskFactors, risk_score: assessed.score, risk_level: effectiveLevel, risk_override: f.riskOverride || null, risk_override_reason: f.riskOverrideReason || null,
      dd_level: ddLevel, documents: f.documents, status: f.status === "not_started" ? "in_progress" : f.status,
      completed_on: f.completedOn || null, completed_by: f.completedBy || null, approved_by: f.approvedBy || null, approved_on: f.approvedOn || null,
      next_review_on: f.nextReviewOn || null, notes: f.notes || null, decision: f.decision || null, decision_reason: f.decisionReason || null, ubo_actions_record: f.uboActionsRecord || null,
      risk_factor_notes: f.riskFactorNotes || {}, assessment_type: f.assessmentType || "initial", assessment_date: f.assessmentDate || toKey(TODAY), sheet_version: f.sheetVersion || 1, risk_score_b: amlAssessB(f.riskFactors).score,
      ubo_change_date: f.uboChangeDate || null, ubo_filed_date: f.uboFiledDate || null,
      updated_at: new Date().toISOString(), ...extra,
    };
    const { error } = await supabase.from("client_aml").upsert(row, { onConflict: "client_id" });
    if (error) { window.alert(`Couldn't save: ${error.message}`); return false; }
    // The representative lives on the client card too — keep both in step.
    await supabase.from("clients").update({ representative_name: f.representativeName || null, representative_afm: f.representativeAfm || null, representative_id_type: f.representativeIdType || null, representative_id_number: f.representativeIdNumber || null }).eq("id", client.id);
    await refetchDueDiligence();
    return true;
  }
  function checklist() {
    const missing = [];
    if (!f.representativeName) missing.push("identification of the client / legal representative");
    if (isLegal && ubos.length === 0 && !f.riskFactors.includes("a_ubo_unknown")) missing.push("at least one beneficial owner");
    if (isLegal && !f.documents.ubo_registry?.collected) missing.push("proof of UBO register entry");
    if (!f.sanctionsCheckedOn) missing.push("PEP / sanctions check date");
    if (!f.relationshipPurpose) missing.push("purpose & nature of the relationship");
    if (!f.infoSources) missing.push("sources of the information");
    if (effectiveLevel === "high" && !f.approvedBy) missing.push("senior approval for high risk");
    if (ddLevel === "enhanced" && !f.sourceOfFunds) missing.push("source of wealth / origin of funds");
    if (needsJustification) missing.push("written justification — Annex II factor without high classification");
    const docsMissing = docsRequired.filter((d) => !f.documents[d.key]?.collected).length;
    if (docsMissing) missing.push(`${docsMissing} required document(s)`);
    return missing;
  }
  async function complete() {
    const missing = checklist();
    if (missing.length && !window.confirm(`Still missing: ${missing.join("; ")}.\n\nComplete the file anyway? (άρ. 14 παρ. 2 allows verification to finish shortly after the initial contact only where the risk is low.)`)) return;
    const review = new Date(TODAY); review.setMonth(review.getMonth() + AML_REVIEW_MONTHS[effectiveLevel]);
    const ok = await save({ status: "completed", decision: "accept", completed_on: toKey(TODAY), completed_by: user.id, next_review_on: f.nextReviewOn || toKey(review), assessment_date: f.assessmentDate || toKey(TODAY) });
    if (!ok) return;
    const open = tasks.filter((t) => t.clientId === client.id && t.kind === "onboarding" && t.stage === "aml" && t.status !== "done");
    if (open.length && window.confirm(`Also mark the ${open.length} due-diligence task(s) of the checklist as done?`)) {
      for (const t of open) await supabase.from("tasks").update({ status: "done", done_at: toKey(TODAY) }).eq("id", t.id);
      refetchTasks();
    }
  }
  async function decline() {
    const reason = window.prompt("Reason for declining the engagement (άρ. 13 παρ. 2 — identification, UBO or purpose could not be completed). This is kept on file, and you should consider whether a report to the Authority is due.", f.decisionReason || "");
    if (reason === null) return;
    await save({ status: "declined", decision: "decline", decision_reason: reason, completed_on: toKey(TODAY), completed_by: user.id });
  }
  async function reopen() { await save({ status: "in_progress", decision: null, completed_on: null, completed_by: null }); }
  // Periodic / extraordinary review: new sheet version, today's date, keeps everything else.
  async function startReview(type) {
    const ok = await save({ status: "in_progress", assessment_type: type, assessment_date: toKey(TODAY), sheet_version: (f.sheetVersion || 1) + 1, completed_on: null, completed_by: null, next_review_on: null });
    if (ok) window.alert(`Έκδοση φύλλου ${(f.sheetVersion || 1) + 1} — ${AML_ASSESSMENT_TYPES[type]} επανεξέταση. Ελέγξτε τους παράγοντες, τα έγγραφα και ολοκληρώστε ξανά τον φάκελο.`);
  }
  // Φύλλο αξιολόγησης κινδύνου — Excel, in the layout of the guide's template.
  function exportRiskSheet() {
    const rows = [];
    AML_AXES.forEach((ax) => {
      rows.push({ "Παράγοντας": ax.label, "Παράρτημα / διάταξη": "", "Βαθμοί (Α)": "", "Βαθμοί (Β)": "", "Συντρέχει;": "", "Μετρούν (Α)": "", "Τεκμηρίωση": "" });
      AML_RISK_FACTORS.filter((rf) => rf.axis === ax.key && !rf.override).forEach((rf) => {
        const on = f.riskFactors.includes(rf.key);
        rows.push({ "Παράγοντας": rf.label, "Παράρτημα / διάταξη": rf.ref, "Βαθμοί (Α)": rf.points, "Βαθμοί (Β)": AML_MODEL_B_POINTS[rf.key] ?? "", "Συντρέχει;": on ? "ΝΑΙ" : "ΟΧΙ", "Μετρούν (Α)": on ? rf.points : 0, "Τεκμηρίωση": f.riskFactorNotes?.[rf.key] || "" });
      });
      rows.push({ "Παράγοντας": "Μερικό σύνολο άξονα", "Μετρούν (Α)": AML_RISK_FACTORS.filter((rf) => rf.axis === ax.key && !rf.override && f.riskFactors.includes(rf.key)).reduce((t, rf) => t + rf.points, 0), "Βαθμοί (Β)": assessedB.axes[ax.key] });
    });
    rows.push({ "Παράγοντας": "ΠΕΡΙΠΤΩΣΕΙΣ «ΥΠΕΡΙΣΧΥΕΙ»" });
    AML_RISK_FACTORS.filter((rf) => rf.override).forEach((rf) => rows.push({ "Παράγοντας": rf.label, "Παράρτημα / διάταξη": rf.ref, "Βαθμοί (Α)": "ΥΠΕΡΙΣΧΥΕΙ", "Συντρέχει;": f.riskFactors.includes(rf.key) ? "ΝΑΙ" : "ΟΧΙ", "Τεκμηρίωση": f.riskFactorNotes?.[rf.key] || "" }));
    exportToExcel(`AML_Φύλλο_Αξιολόγησης_${client.name.replace(/[^\w\u0370-\u03FF]+/g, "_")}_${toKey(TODAY)}.xlsx`, [
      { name: "Στοιχεία", rows: [
        { "Πεδίο": "Πελάτης", "Τιμή": client.name }, { "Πεδίο": "Α.Φ.Μ.", "Τιμή": client.afm || "" }, { "Πεδίο": "Νομική μορφή", "Τιμή": client.legalForm || "" },
        { "Πεδίο": "Έκδοση φύλλου", "Τιμή": f.sheetVersion || 1 }, { "Πεδίο": "Τύπος αξιολόγησης", "Τιμή": AML_ASSESSMENT_TYPES[f.assessmentType || "initial"] }, { "Πεδίο": "Ημερομηνία αξιολόγησης", "Τιμή": f.assessmentDate || "" },
      ] },
      { name: "Αξιολόγηση", rows },
      { name: "Αποτέλεσμα", rows: [
        { "Μέγεθος": "Άθροισμα βαθμών — Μοντέλο Α", "Τιμή": assessed.score }, { "Μέγεθος": "Βαθμός 0-100 — Μοντέλο Β", "Τιμή": assessedB.score },
        { "Μέγεθος": "Περιπτώσεις «ΥΠΕΡΙΣΧΥΕΙ»", "Τιμή": assessed.overrides.length }, { "Μέγεθος": "Παράγοντες Παραρτήματος Ι", "Τιμή": AML_RISK_FACTORS.filter((rf) => f.riskFactors.includes(rf.key) && rf.points < 0).length }, { "Μέγεθος": "Παράγοντες Παραρτήματος ΙΙ", "Τιμή": AML_RISK_FACTORS.filter((rf) => f.riskFactors.includes(rf.key) && rf.points > 0).length },
        { "Μέγεθος": "Κατηγορία κινδύνου", "Τιμή": ({ low: "ΧΑΜΗΛΟΣ", medium: "ΜΕΤΡΙΟΣ", high: "ΥΨΗΛΟΣ" })[effectiveLevel] }, { "Μέγεθος": "Μέτρα δέουσας επιμέλειας", "Τιμή": AML_DD_LABELS[ddLevel] }, { "Μέγεθος": "Ελάχιστη συχνότητα επικαιροποίησης", "Τιμή": effectiveLevel === "low" ? "Ανά διετία" : "Ετησίως και σε κάθε μεταβολή" },
        { "Μέγεθος": "Έλεγχος τεκμηρίου Παραρτήματος ΙΙ", "Τιμή": needsJustification ? "ΑΠΑΙΤΕΙΤΑΙ ΓΡΑΠΤΗ ΑΙΤΙΟΛΟΓΙΑ" : (f.riskOverrideReason || "—") },
        { "Μέγεθος": "Έγκριση ανώτερου στελέχους", "Τιμή": f.approvedBy ? `${employees.find((e) => e.id === f.approvedBy)?.name || ""} ${f.approvedOn || ""}` : (effectiveLevel === "high" ? "ΕΚΚΡΕΜΕΙ" : "Δεν απαιτείται") },
        { "Μέγεθος": "Επόμενη επανεξέταση", "Τιμή": f.nextReviewOn || "" }, { "Μέγεθος": "Υπεύθυνος συμμόρφωσης", "Τιμή": employees.find((e) => e.id === f.completedBy)?.name || "" },
      ] },
    ]);
  }

  function openUbo(x) { setUboModal(x || "new"); setU(x ? { ...x } : { name: "", afm: "", idNumber: "", nationality: "Ελληνική", residenceCountry: "Ελλάδα", ownershipPct: "", controlType: "Ownership", isPep: false, verifiedOn: "", verificationDoc: "", note: "" }); }
  async function saveUbo() {
    if (!u.name?.trim()) { window.alert("Enter the beneficial owner's name."); return; }
    const row = { client_id: client.id, name: u.name.trim(), afm: u.afm || null, id_number: u.idNumber || null, nationality: u.nationality || null, residence_country: u.residenceCountry || null, ownership_pct: u.ownershipPct === "" || u.ownershipPct === null ? null : Number(u.ownershipPct), control_type: u.controlType || null, is_pep: !!u.isPep, verified_on: u.verifiedOn || null, verification_doc: u.verificationDoc || null, note: u.note || null };
    const { error } = uboModal === "new" ? await supabase.from("client_beneficial_owners").insert(row) : await supabase.from("client_beneficial_owners").update(row).eq("id", uboModal.id);
    if (error) { window.alert(`Couldn't save: ${error.message}`); return; }
    if (u.isPep && !f.riskFactors.includes("a_pep")) set({ riskFactors: [...f.riskFactors, "a_pep"], pepStatus: f.pepStatus === "none" ? "ubo" : f.pepStatus });
    await refetchDueDiligence(); setUboModal(null);
  }
  async function deleteUbo(x) { if (!window.confirm(`Remove ${x.name}?`)) return; await supabase.from("client_beneficial_owners").delete().eq("id", x.id); refetchDueDiligence(); }

  const section = (n, title, sub) => (
    <div style={{ marginTop: 20, marginBottom: 10, paddingTop: 14, borderTop: `1px solid ${C.border}` }}>
      <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 13.5, color: C.ink }}><span style={{ fontFamily: mono, color: C.accentDark, marginRight: 8 }}>{n}</span>{title}</div>
      {sub && <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginTop: 2 }}>{sub}</div>}
    </div>
  );
  const grid2 = { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 };
  const missingNow = checklist();

  return (
    <div>
      <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
        <StatCard label="Due diligence file" value={AML_STATUS_LABELS[f.status]} accent={f.status === "completed" ? C.accent : f.status === "declined" ? C.danger : f.status === "in_progress" ? C.warn : undefined} sub={f.completedOn ? `${f.completedOn} · ${empName(f.completedBy)}` : "ν. 4557/2018 · ΠΟΛ.1200/2018"} />
        <StatCard label="Risk classification" value={AML_LEVEL_LABELS[effectiveLevel]} accent={effectiveLevel === "high" ? C.danger : effectiveLevel === "medium" ? C.warn : C.accent} sub={`Sum ${assessed.score >= 0 ? "+" : ""}${assessed.score}${assessed.forced ? " · override case" : ""}${f.riskOverride && !assessed.forced ? " · manually set" : ""}`} />
        <StatCard label="Due diligence level" value={AML_DD_LABELS[ddLevel].split(" ")[0]} sub={AML_DD_LABELS[ddLevel]} />
        <StatCard label="Documents" value={`${docsCollected}/${docsRequired.length}`} sub={`Required for ${AML_ENTITY_TYPES[f.entityType].split(" (")[0]}`} accent={docsCollected < docsRequired.length ? C.warn : C.accent} />
        <StatCard label="Next review" value={f.nextReviewOn || "—"} accent={f.nextReviewOn && fromKey(f.nextReviewOn) < TODAY ? C.danger : undefined} sub={`At least every ${AML_REVIEW_MONTHS[effectiveLevel]} months${effectiveLevel === "high" ? " and on every change" : ""}`} />
      </div>

      <Panel>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted, maxWidth: 640 }}>
            Identification and verification must be completed <strong>before</strong> the relationship starts (άρ. 14 παρ. 1). Keep this file for 5 years after the relationship ends (άρ. 30). Managers only.
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Btn variant="secondary" icon={Receipt} onClick={() => printHtml(`Due diligence — ${client.name}`, buildDueDiligenceHtml({ client, aml: aml ? { ...aml, ...f } : f, ubos, documents, employees, tasks }))}>Download PDF</Btn>
            <Btn variant="secondary" onClick={exportRiskSheet}>Φύλλο αξιολόγησης (Excel)</Btn>
            {(f.status === "completed") && <Btn variant="secondary" onClick={() => startReview("annual")}>Ετήσια επανεξέταση</Btn>}
            {(f.status === "completed") && <Btn variant="ghost" onClick={() => startReview("extraordinary")}>Έκτακτη επανεξέταση</Btn>}
            <Btn variant="secondary" onClick={() => save()} disabled={!dirty}>Save</Btn>
            {f.status !== "completed" && f.status !== "declined" && <Btn variant="ghost" onClick={decline}>Decline engagement</Btn>}
            {f.status !== "completed" && f.status !== "declined" ? <Btn icon={Check} onClick={complete}>Complete file</Btn> : <Btn variant="ghost" onClick={reopen}>Reopen</Btn>}
          </div>
        </div>
        {missingNow.length > 0 && f.status !== "completed" && (
          <div style={{ marginTop: 10, fontFamily: sans, fontSize: 12.5, color: C.warn }}>Open items: {missingNow.join(" · ")}</div>
        )}

        {section("1", "Identification of the client (ΠΟΛ.1200/2018)", "Choose the table that applies; the document list in section 6 follows it. The legal representative is shared with the client's Details tab. If the client acts through an authorised person, that person is identified too.")}
        <div style={grid2}>
          <Field label="Entity type">
            <select style={inputStyle} value={f.entityType} onChange={(e) => set({ entityType: e.target.value })}>
              {Object.entries(AML_ENTITY_TYPES).map(([k, l]) => <option key={k} value={k}>{l}</option>)}
            </select>
          </Field>
          <Field label="ΑΦΜ (from client details)"><input style={{ ...inputStyle, background: C.bg }} value={client.afm || ""} readOnly /></Field>
          {isLegal && <>
            <Field label="ΓΕΜΗ number"><input style={inputStyle} value={f.gemiNumber} onChange={(e) => set({ gemiNumber: e.target.value })} /></Field>
            <Field label="Incorporation date"><input type="date" style={inputStyle} value={f.incorporationDate} onChange={(e) => set({ incorporationDate: e.target.value })} /></Field>
          </>}
          <Field label={isLegal ? "Legal representative — full name & father's name" : "Full name & father's name"}><input style={inputStyle} value={f.representativeName} onChange={(e) => set({ representativeName: e.target.value })} /></Field>
          <Field label="Representative ΑΦΜ"><input style={inputStyle} value={f.representativeAfm} onChange={(e) => set({ representativeAfm: e.target.value })} /></Field>
          <Field label="ID document type">
            <select style={inputStyle} value={f.representativeIdType} onChange={(e) => set({ representativeIdType: e.target.value })}>
              <option value="">—</option><option>Δελτίο ταυτότητας</option><option>Διαβατήριο</option><option>Άδεια διαμονής</option><option>Ταυτότητα ομογενούς</option><option>Υπηρεσιακή ταυτότητα</option>
            </select>
          </Field>
          <Field label="ID number · issuing authority"><input style={inputStyle} value={f.representativeIdNumber} onChange={(e) => set({ representativeIdNumber: e.target.value })} placeholder="e.g. ΑΚ 123456 · Τ.Α. Χαλανδρίου" /></Field>
          <Field label="Authorised person / proxy (if any) — name"><input style={inputStyle} value={f.proxyName} onChange={(e) => set({ proxyName: e.target.value, riskFactors: e.target.value && !f.riskFactors.includes("d_proxy") ? [...f.riskFactors, "d_proxy"] : f.riskFactors })} /></Field>
          <Field label="Proxy ID number & authorisation"><input style={inputStyle} value={f.proxyIdNumber} onChange={(e) => set({ proxyIdNumber: e.target.value })} disabled={!f.proxyName} /></Field>
        </div>

        {isLegal && <>
          {section("2–3", "Beneficial owners & ownership structure", "Natural persons with >25% or control. Proof of entry in the Central UBO Register is required (άρ. 14 παρ. 1) but is never sufficient on its own (άρ. 12 παρ. 4): map the structure from the shareholder register and go up a level wherever >25% belongs to another entity.")}
          <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 8 }}><Btn size="sm" icon={Plus} onClick={() => openUbo(null)}>Add beneficial owner</Btn></div>
          {ubos.length === 0 ? <div style={{ fontFamily: sans, fontSize: 12.5, color: C.warn }}>No beneficial owners recorded yet.</div> : (
            <TableShell headers={["Name", "ΑΦΜ / ID", "Nationality · residence", "Ownership / control", "PEP", "Verified", ""]}>
              {ubos.map((x) => (
                <tr key={x.id}>
                  <Td style={{ fontWeight: 600 }}>{x.name}</Td>
                  <Td mono>{[x.afm, x.idNumber].filter(Boolean).join(" / ") || "—"}</Td>
                  <Td style={{ color: C.inkMuted }}>{[x.nationality, x.residenceCountry].filter(Boolean).join(" · ") || "—"}</Td>
                  <Td mono>{x.ownershipPct !== null ? `${x.ownershipPct}%` : "—"}{x.controlType ? <span style={{ color: C.inkMuted, fontFamily: sans }}> · {x.controlType}</span> : null}</Td>
                  <Td>{x.isPep ? <Badge tone="danger">PEP</Badge> : <span style={{ color: C.inkFaint }}>No</span>}</Td>
                  <Td mono style={{ color: x.verifiedOn ? C.ink : C.warn }}>{x.verifiedOn || "Not verified"}</Td>
                  <Td right><div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}><Btn size="sm" variant="ghost" icon={Pencil} onClick={() => openUbo(x)}>Edit</Btn><Btn size="sm" variant="ghost" icon={Trash2} onClick={() => deleteUbo(x)}>Remove</Btn></div></Td>
                </tr>
              ))}
            </TableShell>
          )}
          <div style={{ marginTop: 10 }}>
            <Field label="If no natural person could be identified: senior managing official treated as UBO — record of actions taken and difficulties (άρ. 3 παρ. 17 αβ΄)">
              <textarea style={textareaStyle} value={f.uboActionsRecord} onChange={(e) => set({ uboActionsRecord: e.target.value, riskFactors: e.target.value && !f.riskFactors.includes("a_ubo_unknown") ? [...f.riskFactors, "a_ubo_unknown"] : f.riskFactors })} placeholder="Leave empty when the beneficial owners above were identified." />
            </Field>
          </div>
        </>}

        {section("4", "PEP and sanctions screening", "Client, legal representative, beneficial owners, and their close relatives and close associates. A PEP forces enhanced due diligence (άρ. 18).")}
        <div style={grid2}>
          <Field label="Politically exposed person?">
            <select style={inputStyle} value={f.pepStatus} onChange={(e) => { const v = e.target.value; set({ pepStatus: v, riskFactors: v === "none" ? f.riskFactors.filter((k) => k !== "a_pep") : (f.riskFactors.includes("a_pep") ? f.riskFactors : [...f.riskFactors, "a_pep"]) }); }}>
              <option value="none">No</option><option value="client">Yes — the client</option><option value="ubo">Yes — a beneficial owner</option><option value="representative">Yes — the representative / close associate</option>
            </select>
          </Field>
          <Field label="PEP details (position, country, relationship)"><input style={inputStyle} value={f.pepDetails} onChange={(e) => set({ pepDetails: e.target.value })} disabled={f.pepStatus === "none"} /></Field>
          <Field label="Lists checked on"><input type="date" style={inputStyle} value={f.sanctionsCheckedOn} onChange={(e) => set({ sanctionsCheckedOn: e.target.value })} /></Field>
          <Field label="Result"><input style={inputStyle} value={f.sanctionsResult} onChange={(e) => set({ sanctionsResult: e.target.value })} placeholder="e.g. No match — EU/UN sanctions, PEP lists" /></Field>
        </div>

        {section("5", "Purpose and nature of the relationship", "Half a page per client: activity, our services, why they came to us, expected transaction profile — and where each piece of information came from (audit question 23).")}
        <div style={grid2}>
          <Field label="Activity / sector (ΚΑΔ, description)"><input style={inputStyle} value={f.activity} onChange={(e) => set({ activity: e.target.value })} /></Field>
          <Field label="Countries (seat, shareholders, UBO, counterparties)"><input style={inputStyle} value={f.countries} onChange={(e) => set({ countries: e.target.value })} /></Field>
          <div style={{ gridColumn: "1 / span 2" }}><Field label="Purpose, why they approached us, expected transaction profile"><textarea style={textareaStyle} value={f.relationshipPurpose} onChange={(e) => set({ relationshipPurpose: e.target.value })} /></Field></div>
          <div style={{ gridColumn: "1 / span 2" }}><Field label="Sources of the information (client declaration, ΓΕΜΗ, ΑΑΔΕ registration, website, prior returns, bank)"><input style={inputStyle} value={f.infoSources} onChange={(e) => set({ infoSources: e.target.value })} /></Field></div>
          <div style={{ gridColumn: "1 / span 2" }}>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 6 }}>Services we provide (an engagement letter describing them covers most of this)</div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {AML_SERVICES.map((sv) => (
                <label key={sv} style={{ display: "flex", alignItems: "center", gap: 6, fontFamily: sans, fontSize: 12.5, color: C.ink, border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 9px", background: f.services.includes(sv) ? C.accentSoft : C.surface, cursor: "pointer" }}>
                  <input type="checkbox" checked={f.services.includes(sv)} onChange={() => toggleService(sv)} />{sv}
                </label>
              ))}
            </div>
          </div>
          <div style={{ gridColumn: "1 / span 2" }}><Field label={`Source of wealth / origin of funds${ddLevel === "enhanced" ? " (required — enhanced due diligence)" : " (optional at this level)"}`}><textarea style={textareaStyle} value={f.sourceOfFunds} onChange={(e) => set({ sourceOfFunds: e.target.value })} /></Field></div>
        </div>

        {section("6", "Risk scoring — Φύλλο αξιολόγησης κινδύνου (άρ. 13 παρ. 9)", "Four axes of άρ. 35 παρ. 1. Model A: Annex I subtracts, Annex II adds; ≤0 with an Annex I factor → low, +1..+5 → medium, ≥+6 or override → high. Model B (0–100) is computed alongside. Write the τεκμηρίωση for every factor you tick.")}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 8 }}>
          <Field label="Τύπος αξιολόγησης"><select style={inputStyle} value={f.assessmentType || "initial"} onChange={(e) => set({ assessmentType: e.target.value })}>{Object.entries(AML_ASSESSMENT_TYPES).map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select></Field>
          <Field label="Ημερομηνία αξιολόγησης"><input type="date" style={inputStyle} value={f.assessmentDate || ""} onChange={(e) => set({ assessmentDate: e.target.value })} /></Field>
          <Field label="Έκδοση φύλλου"><input style={{ ...inputStyle, background: C.bg }} value={f.sheetVersion || 1} readOnly /></Field>
        </div>
        {AML_AXES.map((ax) => (
          <div key={ax.key} style={{ marginBottom: 10 }}>
            <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.inkMuted, margin: "8px 0 4px" }}>{ax.label}</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4 }}>
              {AML_RISK_FACTORS.filter((rf) => rf.axis === ax.key).map((rf) => {
                const on = f.riskFactors.includes(rf.key);
                return (
                  <div key={rf.key} style={{ padding: "4px 8px", borderRadius: 6, background: on ? (rf.override ? "#FDECEC" : C.accentSoft) : "transparent" }}>
                    <label style={{ display: "flex", alignItems: "center", gap: 8, fontFamily: sans, fontSize: 12.5, color: C.ink, cursor: "pointer" }}>
                      <input type="checkbox" checked={on} onChange={() => toggleFactor(rf.key)} />
                      <span style={{ flex: 1 }}>{rf.label}</span>
                      <span style={{ fontFamily: mono, fontSize: 11, color: rf.override ? C.danger : rf.points < 0 ? C.accentDark : C.inkFaint, whiteSpace: "nowrap" }}>{rf.override ? "ΥΠΕΡΙΣΧΥΕΙ" : `${rf.points > 0 ? "+" : ""}${rf.points} · B ${AML_MODEL_B_POINTS[rf.key] ?? 0}`} · {rf.ref}</span>
                    </label>
                    {on && <input style={{ ...inputStyle, marginTop: 4, padding: "4px 8px", fontSize: 12 }} value={f.riskFactorNotes?.[rf.key] || ""} onChange={(e) => setFactorNote(rf.key, e.target.value)} placeholder="Τεκμηρίωση — γιατί συντρέχει, με τα συγκεκριμένα στοιχεία" />}
                  </div>
                );
              })}
            </div>
          </div>
        ))}
        <div style={{ ...grid2, marginTop: 12 }}>
          <div style={{ fontFamily: sans, fontSize: 13, color: C.ink, display: "flex", flexDirection: "column", gap: 6 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              Model A <span style={{ fontFamily: mono, fontWeight: 700 }}>{assessed.score >= 0 ? "+" : ""}{assessed.score}</span>
              · Model B <span style={{ fontFamily: mono, fontWeight: 700 }}>{assessedB.score}/100</span> <span style={{ fontFamily: mono, fontSize: 11, color: C.inkFaint }}>({AML_AXES.map((ax) => `${ax.key} ${assessedB.axes[ax.key]}`).join(" · ")})</span>
              {assessed.forced && <Badge tone="danger">Υπερισχύει: {assessed.overrides.map((o) => o.ref).join(", ")}</Badge>}
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              → <Badge tone={AML_LEVEL_TONE[effectiveLevel]}>{AML_LEVEL_LABELS[effectiveLevel]} risk</Badge> <Badge tone={AML_LEVEL_TONE[effectiveLevel]}>{AML_DD_LABELS[ddLevel]}</Badge>
              {assessedB.level !== assessed.level && !assessed.forced && <span style={{ fontFamily: sans, fontSize: 11.5, color: C.warn }}>Model B would give {AML_LEVEL_LABELS[assessedB.level]} — the firm's policy applies one model consistently.</span>}
              {needsJustification && <span style={{ fontFamily: sans, fontSize: 11.5, color: C.danger }}>Annex II factor present without high classification — written justification required below.</span>}
            </div>
          </div>
          <Field label={needsJustification ? "Γραπτή αιτιολογία (required): γιατί ο παράγοντας Παρ. ΙΙ δεν ανεβάζει τον κίνδυνο" : "Documented deviation from the computed level (optional, reason required)"}>
            <div style={{ display: "flex", gap: 8 }}>
              <select style={{ ...inputStyle, width: 130 }} value={f.riskOverride} onChange={(e) => set({ riskOverride: e.target.value })} disabled={assessed.forced}>
                <option value="">Computed</option>{Object.entries(AML_LEVEL_LABELS).map(([k, l]) => <option key={k} value={k}>{l}</option>)}
              </select>
              <input style={{ ...inputStyle, ...(needsJustification ? { border: `1px solid ${C.danger}` } : {}) }} value={f.riskOverrideReason} onChange={(e) => set({ riskOverrideReason: e.target.value })} placeholder="Why the factor does not raise the risk in this case" disabled={assessed.forced || (!f.riskOverride && !annexIIPresent)} />
            </div>
          </Field>
          {effectiveLevel === "high" && <>
            <Field label="High risk — approval of senior management before engagement">
              <select style={inputStyle} value={f.approvedBy || ""} onChange={(e) => set({ approvedBy: e.target.value || null, approvedOn: e.target.value ? (f.approvedOn || toKey(TODAY)) : "" })}>
                <option value="">— not approved yet —</option>
                {employees.filter((e) => e.role === "admin").map((e) => <option key={e.id} value={e.id}>{e.name}</option>)}
              </select>
            </Field>
            <Field label="Approved on"><input type="date" style={inputStyle} value={f.approvedOn} onChange={(e) => set({ approvedOn: e.target.value })} /></Field>
          </>}
        </div>

        {section("7", `Documents on file — ${AML_ENTITY_TYPES[f.entityType]}`, `${docsCollected} of ${docsRequired.length} required collected`)}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
          {docsForType.map((d) => {
            const st = f.documents[d.key] || {};
            const needed = !d.optional && (!d.enhancedOnly || ddLevel === "enhanced");
            const files = documents.filter((x) => x.category === d.key);
            return (
              <div key={d.key} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 8px", borderRadius: 6, background: st.collected ? C.accentSoft : "transparent", opacity: needed ? 1 : 0.6, flexWrap: "wrap" }}>
                <input type="checkbox" checked={!!st.collected} onChange={(e) => setDoc(d.key, { collected: e.target.checked, date: e.target.checked ? (st.date || toKey(TODAY)) : "" })} />
                <span style={{ flex: 1, fontFamily: sans, fontSize: 12.5, color: C.ink, minWidth: 220 }}>{d.label}{!needed ? " — optional" : ""}
                  {files.map((x) => <div key={x.id} style={{ fontSize: 11.5 }}><button onClick={() => openClientDocument(x)} style={{ ...iconBtnStyle, padding: 0, fontFamily: sans, fontSize: 11.5, color: C.accentDark, textDecoration: "underline" }}>{x.fileName}</button> <span style={{ color: C.inkFaint }}>{x.uploadedAt.slice(0, 10)}</span></div>)}
                </span>
                <input type="date" style={{ ...inputStyle, width: 140, padding: "4px 6px", fontSize: 12 }} value={st.date || ""} onChange={(e) => setDoc(d.key, { date: e.target.value })} disabled={!st.collected} />
                {d.key === "kyc_form" && <Btn variant="secondary" icon={Receipt} onClick={() => printHtml(`KYC — ${client.name}`, buildKycQuestionnaireHtml({ client, aml: { ...(aml || {}), ...f }, ubos }), { editable: true })}>Print KYC form</Btn>}
                <UploadButton label={files.length ? "Upload another" : "Upload"} onFile={async (file) => { if (await uploadClientDocument(client, file, d.key, "", user)) { setDoc(d.key, { collected: true, date: st.date || toKey(TODAY) }); refetchDocuments && refetchDocuments(); } }} />
              </div>
            );
          })}
        </div>

        {section("8", "Decision, file and next review", "Accepting opens the file and sets the review date; declining is recorded with its reason (άρ. 13 παρ. 2).")}
        <div style={grid2}>
          <Field label="Μεταβολή πραγματικών δικαιούχων διαπιστώθηκε στις (ενεργοποιεί έκτακτη επανεξέταση & 60-ήμερη προθεσμία ΚΜΠΔ, άρ. 20)"><input type="date" style={inputStyle} value={f.uboChangeDate || ""} onChange={(e) => set({ uboChangeDate: e.target.value })} /></Field>
          <Field label="Ο πελάτης καταχώρισε τη μεταβολή στο ΚΜΠΔ στις"><input type="date" style={inputStyle} value={f.uboFiledDate || ""} onChange={(e) => set({ uboFiledDate: e.target.value })} disabled={!f.uboChangeDate} /></Field>
          <Field label="Next periodic review"><input type="date" style={inputStyle} value={f.nextReviewOn} onChange={(e) => set({ nextReviewOn: e.target.value })} /></Field>
          <Field label="Status"><div style={{ paddingTop: 7 }}><Badge tone={AML_STATUS_TONE[f.status]}>{AML_STATUS_LABELS[f.status]}</Badge>{f.completedOn ? <span style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginLeft: 8 }}>{f.completedOn} · {empName(f.completedBy)}</span> : null}</div></Field>
          {f.status === "declined" && <div style={{ gridColumn: "1 / span 2", fontFamily: sans, fontSize: 12.5, color: C.danger }}>Declined: {f.decisionReason || "—"}</div>}
          <div style={{ gridColumn: "1 / span 2" }}><Field label="Notes / conclusions of the compliance officer"><textarea style={textareaStyle} value={f.notes} onChange={(e) => set({ notes: e.target.value })} /></Field></div>
        </div>
        <div style={{ display: "flex", gap: 8, marginTop: 14, alignItems: "center" }}>
          <Btn onClick={() => save()} disabled={!dirty}>Save</Btn>
          {dirty && <span style={{ fontFamily: sans, fontSize: 12, color: C.warn }}>Unsaved changes</span>}
        </div>
      </Panel>

      {uboModal && (
        <Modal title={uboModal === "new" ? "Add beneficial owner" : "Edit beneficial owner"} onClose={() => setUboModal(null)} width={460}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Field label="Full name & father's name"><input style={inputStyle} value={u.name} onChange={(e) => setU({ ...u, name: e.target.value })} autoFocus /></Field>
            <div style={grid2}>
              <Field label="ΑΦΜ"><input style={inputStyle} value={u.afm} onChange={(e) => setU({ ...u, afm: e.target.value })} /></Field>
              <Field label="ID / passport number"><input style={inputStyle} value={u.idNumber} onChange={(e) => setU({ ...u, idNumber: e.target.value })} /></Field>
              <Field label="Nationality"><input style={inputStyle} value={u.nationality} onChange={(e) => setU({ ...u, nationality: e.target.value })} /></Field>
              <Field label="Country of residence"><input style={inputStyle} value={u.residenceCountry} onChange={(e) => setU({ ...u, residenceCountry: e.target.value })} /></Field>
              <Field label="Ownership %"><input type="number" min="0" max="100" style={inputStyle} value={u.ownershipPct ?? ""} onChange={(e) => setU({ ...u, ownershipPct: e.target.value })} /></Field>
              <Field label="Type of control">
                <select style={inputStyle} value={u.controlType} onChange={(e) => setU({ ...u, controlType: e.target.value })}>
                  <option>Ownership</option><option>Voting rights</option><option>Indirect (via entity)</option><option>Senior managing official</option><option>Trust — settlor/trustee/protector/beneficiary</option><option>Other control</option>
                </select>
              </Field>
              <Field label="Verified on"><input type="date" style={inputStyle} value={u.verifiedOn} onChange={(e) => setU({ ...u, verifiedOn: e.target.value })} /></Field>
              <Field label="Verification document"><input style={inputStyle} value={u.verificationDoc} onChange={(e) => setU({ ...u, verificationDoc: e.target.value })} placeholder="e.g. ID copy + ΚΜΠΔ extract 10/2026" /></Field>
            </div>
            <label style={{ display: "flex", alignItems: "center", gap: 8, fontFamily: sans, fontSize: 13, color: C.ink }}><input type="checkbox" checked={!!u.isPep} onChange={(e) => setU({ ...u, isPep: e.target.checked })} /> PEP, close relative or close associate</label>
            <Field label="Note"><input style={inputStyle} value={u.note} onChange={(e) => setU({ ...u, note: e.target.value })} /></Field>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}><Btn variant="secondary" onClick={() => setUboModal(null)}>Cancel</Btn><Btn onClick={saveUbo}>Save</Btn></div>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Client card — profile, contacts, notes & communication, tasks           */
/* ---------------------------------------------------------------------- */

const formFromClient = (c) => ({ name: c.name, afm: c.afm, doy: c.doy, legalForm: c.legalForm, address: c.address, email: c.email, phone: c.phone, accountManagerId: c.accountManagerId || "", generalNotes: c.generalNotes,
  representativeName: c.representativeName, representativeAfm: c.representativeAfm, representativeIdType: c.representativeIdType, representativeIdNumber: c.representativeIdNumber, representativeEmail: c.representativeEmail, representativePhone: c.representativePhone });
function ClientCard({ user, client, employees, contacts, notes, tasks, entries, templates = [], aml = null, ubos = [], refetchDueDiligence, documents = [], refetchDocuments, firm = DEFAULT_FIRM, saveFirm, initialTab = "details", refetchClients, refetchContacts, refetchNotes, refetchTasks, onBack }) {
  const [showFlow, setShowFlow] = useState(false);
  const [tab, setTab] = useState(initialTab);
  useEffect(() => { setTab(initialTab); }, [initialTab, client?.id]);
  const [form, setForm] = useState(null);
  const [contactModal, setContactModal] = useState(null); // null | "new" | contact
  const [cName, setCName] = useState(""); const [cRole, setCRole] = useState(""); const [cEmail, setCEmail] = useState(""); const [cPhone, setCPhone] = useState(""); const [cNote, setCNote] = useState(""); const [cPrimary, setCPrimary] = useState(false);
  const [noteModal, setNoteModal] = useState(null); // null | "new" | note
  const [taskModal, setTaskModal] = useState(null);

  useEffect(() => {
    if (!client) return;
    setForm(formFromClient(client));
  }, [client?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  if (!client || !form) return <div><Btn variant="ghost" icon={ArrowLeft} onClick={onBack}>Back to clients</Btn><EmptyState icon={Building2} title="Client not found" /></div>;

  const myContacts = contacts.filter((x) => x.clientId === client.id).sort((a, b) => (b.isPrimary ? 1 : 0) - (a.isPrimary ? 1 : 0) || a.name.localeCompare(b.name));
  const myNotes = notes.filter((n) => n.clientId === client.id).sort((a, b) => b.noteDate.localeCompare(a.noteDate) || b.createdAt.localeCompare(a.createdAt));
  const myTasks = tasks.filter((t) => t.clientId === client.id);
  const openTasks = myTasks.filter((t) => t.status !== "done");
  const empName = (id) => employees.find((e) => e.id === id)?.name || "—";
  const dirty = JSON.stringify(form) !== JSON.stringify(formFromClient(client));

  // Quick activity figures for the header: hours & revenue year to date.
  const yearStart = new Date(TODAY.getFullYear(), 0, 1);
  const isAdminUser = user.role === "admin";
  const ytdEntries = isAdminUser ? entries.filter((e) => e.clientId === client.id && fromKey(e.date) >= yearStart && fromKey(e.date) <= TODAY) : [];
  const ytdHours = ytdEntries.reduce((s, e) => s + e.hours, 0);
  const ytdRevenue = periodClientMetrics(client, ytdEntries, yearStart, TODAY).revenue;

  async function saveDetails() {
    if (!form.name.trim()) { window.alert("The client needs a name."); return; }
    const { error } = await supabase.from("clients").update({
      name: form.name.trim(), afm: form.afm.trim() || null, doy: form.doy.trim() || null, legal_form: form.legalForm || null,
      address: form.address.trim() || null, email: form.email.trim() || null, phone: form.phone.trim() || null,
      account_manager_id: form.accountManagerId || null, general_notes: form.generalNotes.trim() || null,
      representative_name: form.representativeName.trim() || null, representative_afm: form.representativeAfm.trim() || null, representative_id_type: form.representativeIdType || null, representative_id_number: form.representativeIdNumber.trim() || null,
      representative_email: form.representativeEmail.trim() || null, representative_phone: form.representativePhone.trim() || null,
    }).eq("id", client.id);
    if (error) { window.alert(`Couldn't save: ${error.message}`); return; }
    // Keep the due-diligence record in step: it shows the same person.
    if (aml) await supabase.from("client_aml").update({ representative_name: form.representativeName.trim() || null, representative_afm: form.representativeAfm.trim() || null, representative_id_type: form.representativeIdType || null, representative_id_number: form.representativeIdNumber.trim() || null }).eq("client_id", client.id);
    refetchClients(); if (aml && refetchDueDiligence) refetchDueDiligence();
  }
  function openContact(x) {
    setContactModal(x || "new");
    setCName(x?.name || ""); setCRole(x?.role || ""); setCEmail(x?.email || ""); setCPhone(x?.phone || ""); setCNote(x?.note || ""); setCPrimary(x ? x.isPrimary : myContacts.length === 0);
  }
  async function saveContact() {
    if (!cName.trim()) { window.alert("Enter the contact's name."); return; }
    const row = { client_id: client.id, name: cName.trim(), role: cRole.trim() || null, email: cEmail.trim() || null, phone: cPhone.trim() || null, note: cNote.trim() || null, is_primary: cPrimary };
    if (cPrimary) await supabase.from("client_contacts").update({ is_primary: false }).eq("client_id", client.id);
    const { error } = contactModal === "new"
      ? await supabase.from("client_contacts").insert(row)
      : await supabase.from("client_contacts").update(row).eq("id", contactModal.id);
    if (error) { window.alert(`Couldn't save contact: ${error.message}`); return; }
    await refetchContacts();
    setContactModal(null);
  }
  async function deleteContact(x) {
    if (!window.confirm(`Remove ${x.name}?`)) return;
    await supabase.from("client_contacts").delete().eq("id", x.id);
    refetchContacts();
  }
  async function quickStatus(t, status) {
    await supabase.from("tasks").update({ status, done_at: status === "done" ? toKey(TODAY) : null }).eq("id", t.id);
    refetchTasks();
  }

  const progress = onboardingProgress(client, tasks);
  async function setLifecycle(lifecycle) {
    const patch = { lifecycle, go_live_date: lifecycle === "live" ? (client.goLiveDate || toKey(TODAY)) : client.goLiveDate };
    const { error } = await supabase.from("clients").update(patch).eq("id", client.id);
    if (error) { window.alert(`Couldn't update: ${error.message}`); return; }
    refetchClients();
  }
  async function beginOnboarding() {
    const includeFormation = window.confirm("Is this a newly formed company?\n\nOK = also create the company formation tasks (ΥΜΣ, μισθωτήριο έδρας, τράπεζα, ΕΦΚΑ, κεφάλαιο, ΓΕΜΗ…).\nCancel = existing company, setup tasks only.");
    if (await startOnboarding(client, templates, tasks, user, { includeFormation })) { await refetchTasks(); await refetchClients(); setTab("tasks"); }
  }
  async function addFormationTasks() {
    if (await startOnboarding(client, templates, tasks, user, { includeFormation: true })) { await refetchTasks(); await refetchClients(); setTab("tasks"); }
  }
  async function markGoLive() {
    const gl = goLiveReadiness(client, tasks, aml, documents);
    if (!gl.ready && !window.confirm(`Still missing: ${gl.checks.filter((c) => !c.ok).map((c) => c.label).join("; ")}.\n\nMark the client as live anyway?`)) return;
    await setLifecycle("live");
  }

  const tabs = [
    { key: "details", label: "Details" },
    { key: "fees", label: "Fees & hours" },
    { key: "agreement", label: `Agreement · ${CONTRACT_STATUS_LABELS[client.contractStatus]}` },
    { key: "dd", label: `Due diligence${aml ? ` · ${AML_STATUS_LABELS[aml.status]}` : ""}` },
    { key: "documents", label: `Documents (${documents.length})` },
    { key: "contacts", label: `Contacts (${myContacts.length})` },
    { key: "notes", label: `Notes (${myNotes.length})` },
    { key: "tasks", label: `Tasks (${openTasks.length} open)` },
  ].filter((t) => isAdminUser || !["agreement", "dd", "documents"].includes(t.key));
  const primary = myContacts.find((x) => x.isPrimary);

  return (
    <div>
      <div style={{ marginBottom: 10 }}><Btn variant="ghost" size="sm" icon={ArrowLeft} onClick={onBack}>Back to clients</Btn></div>
      <PageHeader title={<span style={{ display: "inline-flex", alignItems: "center", gap: 10 }}>{client.name}<Badge tone={LIFECYCLE_TONE[client.lifecycle]}>{LIFECYCLE_LABELS[client.lifecycle]}</Badge></span>}
        sub={[client.legalForm, client.afm ? `ΑΦΜ ${client.afm}` : null, client.doy ? `ΔΟΥ ${client.doy}` : null, client.representativeName ? `Rep.: ${client.representativeName}` : null, client.accountManagerId ? `Account manager: ${empName(client.accountManagerId)}` : null, client.goLiveDate ? `Go-live ${client.goLiveDate}` : null].filter(Boolean).join(" · ") || "No profile details yet — fill in the Details tab."}
        right={
          <div style={{ display: "flex", gap: 8 }}>
            <Btn variant="secondary" icon={Euro} onClick={() => setTab("fees")}>Add extra fee</Btn>
            {client.lifecycle === "lead" && <Btn variant="secondary" icon={ListChecks} onClick={beginOnboarding}>Start onboarding</Btn>}
            {client.lifecycle === "onboarding" && <Btn variant="secondary" icon={Check} onClick={markGoLive}>Mark go-live</Btn>}
            <Btn variant="secondary" icon={MessageSquare} onClick={() => setNoteModal("new")}>Add note</Btn>
            <Btn icon={Plus} onClick={() => setTaskModal("new")}>New task</Btn>
          </div>
        } />

      {client.lifecycle === "onboarding" && (
        <Panel style={{ marginBottom: 18, borderLeft: `4px solid ${C.warn}` }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
            <div>
              <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink }}>Onboarding in progress — {progress.done} of {progress.total} setup tasks done</div>
              <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginTop: 2 }}>
                {progress.total === 0 ? "No checklist tasks exist for this client yet." : progress.complete ? "All tasks done." : "The client goes live automatically when all four conditions below are met."}
              </div>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <Btn size="sm" variant="ghost" onClick={() => setShowFlow(true)}>ⓘ Process</Btn>
              {progress.total === 0 && <Btn size="sm" variant="secondary" onClick={beginOnboarding}>Create checklist tasks</Btn>}
              {progress.total > 0 && !progress.list.some((t) => t.stage === "formation") && <Btn size="sm" variant="secondary" onClick={addFormationTasks}>Add formation tasks</Btn>}
              <Btn size="sm" variant="ghost" onClick={() => setTab("tasks")}>Open tasks</Btn>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
            {goLiveReadiness(client, tasks, aml, documents).checks.map((c) => <Badge key={c.key} tone={c.ok ? "accent" : "warn"}>{c.ok ? "✓" : "○"} {c.label}</Badge>)}
          </div>
          {progress.total > 0 && (
            <div style={{ display: "flex", gap: 16, marginTop: 10, flexWrap: "wrap" }}>
              {ONBOARDING_STAGES.map((stg) => {
                const list = progress.list.filter((t) => (t.stage || "setup") === stg.key);
                if (list.length === 0) return null;
                const done = list.filter((t) => t.status === "done").length;
                return (
                  <div key={stg.key} style={{ flex: 1, minWidth: 180 }}>
                    <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 4 }}>{stg.label} <span style={{ fontFamily: mono }}>{done}/{list.length}</span></div>
                    <div style={{ height: 8, background: C.border, borderRadius: 4, overflow: "hidden" }}><div style={{ width: `${(done / list.length) * 100}%`, height: "100%", background: done === list.length ? C.accent : C.warn }} /></div>
                  </div>
                );
              })}
            </div>
          )}
        </Panel>
      )}
      {showFlow && <Modal title="Client onboarding process" onClose={() => setShowFlow(false)} width={680}><OnboardingFlowchart /></Modal>}
      {client.lifecycle === "lead" && (
        <Panel style={{ marginBottom: 18, borderLeft: `4px solid ${C.inkFaint}` }}>
          <Btn size="sm" variant="ghost" onClick={() => setShowFlow(true)}>ⓘ Process</Btn>
          <div style={{ fontFamily: sans, fontSize: 13, color: C.ink }}>Prospect — not yet set up. Press <strong>Start onboarding</strong> when the engagement is agreed: the setup checklist ({templates.filter((t) => t.active).length} tasks) is created for this client and assigned to the account manager.</div>
        </Panel>
      )}

      <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
        <StatCard label="Status" value={client.endDate && client.endDate <= toKey(TODAY) ? "Ended" : client.active ? "Active" : "Inactive"} sub={client.endDate ? `End date ${client.endDate}` : `Fixed fee ${client.fixedFee ? fmtEur(client.fixedFee) + "/mo" : "—"}`} />
        {isAdminUser && <StatCard label={`Hours ${TODAY.getFullYear()} YTD`} value={ytdHours.toFixed(1) + "h"} sub="Logged by the team" />}
        <StatCard label={`Revenue ${TODAY.getFullYear()} YTD`} value={fmtEur(ytdRevenue)} accent={C.accent} sub="Fixed + extra fees" />
        <StatCard label="Open tasks" value={String(openTasks.length)} accent={openTasks.some(taskIsOverdue) ? C.danger : undefined} sub={openTasks.filter(taskIsOverdue).length ? `${openTasks.filter(taskIsOverdue).length} overdue` : "None overdue"} />
        {primary && <StatCard label="Primary contact" value={primary.name} sub={[primary.phone, primary.email].filter(Boolean).join(" · ") || primary.role || ""} />}
      </div>

      <div style={{ display: "flex", gap: 4, borderBottom: `1px solid ${C.border}`, marginBottom: 16 }}>
        {tabs.map((t) => (
          <button key={t.key} onClick={() => setTab(t.key)} style={{ background: "none", border: "none", cursor: "pointer", padding: "8px 14px", fontFamily: sans, fontSize: 13, fontWeight: tab === t.key ? 700 : 500, color: tab === t.key ? C.accentDark : C.inkMuted, borderBottom: tab === t.key ? `2px solid ${C.accent}` : "2px solid transparent", marginBottom: -1 }}>
            {t.label}
          </button>
        ))}
      </div>

      {tab === "details" && (
        <Panel>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, maxWidth: 760 }}>
            <Field label="Client name"><input style={inputStyle} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
            <Field label="Legal form">
              <select style={inputStyle} value={form.legalForm} onChange={(e) => setForm({ ...form, legalForm: e.target.value })}>
                <option value="">—</option>
                {LEGAL_FORMS.map((l) => <option key={l} value={l}>{l}</option>)}
              </select>
            </Field>
            <Field label="ΑΦΜ"><input style={inputStyle} value={form.afm} onChange={(e) => setForm({ ...form, afm: e.target.value })} placeholder="9 digits" /></Field>
            <Field label="ΔΟΥ"><input style={inputStyle} value={form.doy} onChange={(e) => setForm({ ...form, doy: e.target.value })} placeholder="e.g. Α' Αθηνών" /></Field>
            <div style={{ gridColumn: "1 / span 2" }}><Field label="Address"><input style={inputStyle} value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} /></Field></div>
            <Field label="Email"><input style={inputStyle} value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} /></Field>
            <Field label="Phone"><input style={inputStyle} value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
            <Field label="Account manager (at the firm)">
              <select style={inputStyle} value={form.accountManagerId} onChange={(e) => setForm({ ...form, accountManagerId: e.target.value })}>
                <option value="">—</option>
                {employees.filter((e) => e.active || e.id === form.accountManagerId).map((e) => <option key={e.id} value={e.id}>{e.name}</option>)}
              </select>
            </Field>
            <div style={{ gridColumn: "1 / span 2", fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: C.ink, marginTop: 6, paddingTop: 10, borderTop: `1px solid ${C.border}` }}>Legal representative <span style={{ fontWeight: 400, color: C.inkMuted }}>— used in the agreement and the due-diligence file</span></div>
            <Field label="Full name & father's name"><input style={inputStyle} value={form.representativeName} onChange={(e) => setForm({ ...form, representativeName: e.target.value })} /></Field>
            <Field label="ΑΦΜ"><input style={inputStyle} value={form.representativeAfm} onChange={(e) => setForm({ ...form, representativeAfm: e.target.value })} /></Field>
            <Field label="ID document">
              <select style={inputStyle} value={form.representativeIdType} onChange={(e) => setForm({ ...form, representativeIdType: e.target.value })}>
                <option value="">—</option><option>Δελτίο ταυτότητας</option><option>Διαβατήριο</option><option>Άδεια διαμονής</option><option>Ταυτότητα ομογενούς</option><option>Υπηρεσιακή ταυτότητα</option>
              </select>
            </Field>
            <Field label="ID number (ΑΔΤ) · issuing authority"><input style={inputStyle} value={form.representativeIdNumber} onChange={(e) => setForm({ ...form, representativeIdNumber: e.target.value })} placeholder="e.g. ΑΚ 123456" /></Field>
            <Field label="Email"><input style={inputStyle} value={form.representativeEmail} onChange={(e) => setForm({ ...form, representativeEmail: e.target.value })} /></Field>
            <Field label="Phone"><input style={inputStyle} value={form.representativePhone} onChange={(e) => setForm({ ...form, representativePhone: e.target.value })} /></Field>
            <Field label="Stage">
              <select style={inputStyle} value={client.lifecycle} onChange={(e) => setLifecycle(e.target.value)}>
                {Object.entries(LIFECYCLE_LABELS).map(([k, l]) => <option key={k} value={k}>{l}</option>)}
              </select>
            </Field>
            <Field label="Engagement / contract date (onboarding deadlines)">
              <input type="date" style={inputStyle} value={client.contractSignedOn || client.engagementDate || ""} disabled={!!client.contractSignedOn} title={client.contractSignedOn ? "Taken from the signed agreement" : ""}
                onChange={async (e) => { const v = e.target.value || null; await supabase.from("clients").update({ engagement_date: v }).eq("id", client.id); await applyOnboardingDeadlines(client.id, v); refetchClients(); refetchTasks(); }} />
            </Field>
            <Field label="Go-live date">
              <input type="date" style={inputStyle} value={client.goLiveDate || ""} onChange={async (e) => { await supabase.from("clients").update({ go_live_date: e.target.value || null }).eq("id", client.id); refetchClients(); }} />
            </Field>
            <div style={{ gridColumn: "1 / span 2" }}><Field label="General notes (standing information about the client)"><textarea style={textareaStyle} value={form.generalNotes} onChange={(e) => setForm({ ...form, generalNotes: e.target.value })} /></Field></div>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 14, alignItems: "center" }}>
            <Btn onClick={saveDetails} disabled={!dirty}>Save details</Btn>
            {dirty && <span style={{ fontFamily: sans, fontSize: 12, color: C.warn }}>Unsaved changes</span>}
            <span style={{ fontFamily: sans, fontSize: 12, color: C.inkFaint, marginLeft: "auto" }}>Fees, extra fees and hours are in the "Fees & hours" tab.</span>
          </div>
        </Panel>
      )}

      {tab === "contacts" && (
        <Panel>
          <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 10 }}><Btn size="sm" icon={Plus} onClick={() => openContact(null)}>Add contact</Btn></div>
          {myContacts.length === 0 ? <EmptyState icon={UserRound} title="No contact persons yet" /> : (
            <TableShell headers={["Name", "Role", "Phone", "Email", "Note", ""]}>
              {myContacts.map((x) => (
                <tr key={x.id}>
                  <Td><span style={{ fontWeight: 600 }}>{x.name}</span>{x.isPrimary && <Badge tone="accent">Primary</Badge>}</Td>
                  <Td style={{ color: C.inkMuted }}>{x.role || "—"}</Td>
                  <Td mono>{x.phone ? <a href={`tel:${x.phone}`} style={{ color: C.ink, textDecoration: "none" }}><Phone size={11} style={{ marginRight: 4 }} />{x.phone}</a> : "—"}</Td>
                  <Td>{x.email ? <a href={`mailto:${x.email}`} style={{ color: C.accentDark }}><Mail size={11} style={{ marginRight: 4 }} />{x.email}</a> : "—"}</Td>
                  <Td style={{ color: C.inkMuted }}>{x.note || "—"}</Td>
                  <Td right>
                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      <Btn size="sm" variant="ghost" icon={Pencil} onClick={() => openContact(x)}>Edit</Btn>
                      <Btn size="sm" variant="ghost" icon={Trash2} onClick={() => deleteContact(x)}>Remove</Btn>
                    </div>
                  </Td>
                </tr>
              ))}
            </TableShell>
          )}
        </Panel>
      )}

      {tab === "notes" && (
        <Panel>
          {myNotes.length === 0 ? <EmptyState icon={MessageSquare} title="No notes yet" sub="Log calls, emails and meetings so the whole team has the history." /> : (
            <div>
              {myNotes.map((n) => (
                <div key={n.id} style={{ display: "flex", gap: 12, padding: "10px 0", borderBottom: `1px solid ${C.border}`, fontFamily: sans }}>
                  <div style={{ minWidth: 92, fontFamily: mono, fontSize: 12, color: C.inkMuted }}>{n.noteDate}</div>
                  <div style={{ minWidth: 72 }}><Badge tone="neutral">{NOTE_KINDS[n.kind]}</Badge></div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 13, color: C.ink, whiteSpace: "pre-wrap" }}>{n.content}</div>
                    <div style={{ fontSize: 11, color: C.inkFaint, marginTop: 3 }}>{empName(n.authorId)}</div>
                  </div>
                  {(n.authorId === user.id || user.role === "admin") && <button onClick={() => setNoteModal(n)} style={iconBtnStyle} title="Edit"><Pencil size={13} color={C.inkMuted} /></button>}
                </div>
              ))}
            </div>
          )}
        </Panel>
      )}

      {tab === "fees" && <ClientFeesTab client={client} entries={isAdminUser ? entries : []} employees={employees} refetchClients={refetchClients} isAdmin={isAdminUser} />}
      {tab === "agreement" && <ClientAgreementTab user={user} client={client} aml={aml} firm={firm} saveFirm={saveFirm} employees={employees} contacts={myContacts} ubos={ubos} documents={documents} refetchClients={refetchClients} refetchDocuments={refetchDocuments} />}
      {tab === "documents" && <ClientDocumentsTab user={user} client={client} documents={documents} employees={employees} refetchDocuments={refetchDocuments} aml={aml} ubos={ubos} contacts={myContacts} firm={firm} />}
      {tab === "dd" && (
        <DueDiligenceTab user={user} client={client} employees={employees} aml={aml} ubos={ubos} tasks={tasks} documents={documents} refetchDocuments={refetchDocuments} refetchDueDiligence={refetchDueDiligence} refetchTasks={refetchTasks} />
      )}

      {tab === "tasks" && (
        <div>
          {progress.total > 0 && ONBOARDING_STAGES.map((stg) => {
            const list = progress.list.filter((t) => (t.stage || "setup") === stg.key);
            if (list.length === 0) return null;
            const done = list.filter((t) => t.status === "done").length;
            return (
              <Panel key={stg.key} style={{ marginBottom: 18, borderLeft: `4px solid ${done === list.length ? C.accent : C.warn}` }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                  <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink }}>{stg.label} <span style={{ fontFamily: mono, fontSize: 12, color: C.inkMuted }}>{done}/{list.length}</span></div>
                  {stg.key === "aml" && <Btn size="sm" variant="ghost" onClick={() => setTab("dd")}>Open due diligence record</Btn>}
                </div>
                <TaskList user={user} tasks={list} clients={[]} employees={employees} showClient={false} onEdit={(t) => setTaskModal(t)} onQuickStatus={quickStatus} />
              </Panel>
            );
          })}
          <Panel>
            <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 10 }}>Other tasks</div>
            <TaskList user={user} tasks={myTasks.filter((t) => t.kind !== "onboarding")} clients={[]} employees={employees} showClient={false} onEdit={(t) => setTaskModal(t)} onQuickStatus={quickStatus} />
          </Panel>
        </div>
      )}

      {contactModal && (
        <Modal title={contactModal === "new" ? "Add contact" : "Edit contact"} onClose={() => setContactModal(null)} width={440}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Field label="Name"><input style={inputStyle} value={cName} onChange={(e) => setCName(e.target.value)} autoFocus /></Field>
            <Field label="Role / position"><input style={inputStyle} value={cRole} onChange={(e) => setCRole(e.target.value)} placeholder="e.g. Διαχειριστής, Λογιστήριο" /></Field>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <Field label="Phone"><input style={inputStyle} value={cPhone} onChange={(e) => setCPhone(e.target.value)} /></Field>
              <Field label="Email"><input style={inputStyle} value={cEmail} onChange={(e) => setCEmail(e.target.value)} /></Field>
            </div>
            <Field label="Note (optional)"><input style={inputStyle} value={cNote} onChange={(e) => setCNote(e.target.value)} /></Field>
            <label style={{ display: "flex", alignItems: "center", gap: 8, fontFamily: sans, fontSize: 13, color: C.ink }}>
              <input type="checkbox" checked={cPrimary} onChange={(e) => setCPrimary(e.target.checked)} /> Primary contact
            </label>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 6 }}>
              <Btn variant="secondary" onClick={() => setContactModal(null)}>Cancel</Btn>
              <Btn onClick={saveContact}>Save</Btn>
            </div>
          </div>
        </Modal>
      )}
      {noteModal && <NoteModal user={user} note={noteModal === "new" ? null : noteModal} fixedClientId={client.id} clients={[]} onClose={() => setNoteModal(null)} onSaved={refetchNotes} />}
      {taskModal && <TaskModal user={user} task={taskModal === "new" ? null : taskModal} fixedClientId={client.id} clients={[]} employees={employees} onClose={() => setTaskModal(null)} onSaved={refetchTasks} />}
    </div>
  );
}
/* ---------------------------------------------------------------------- */
/* AML Compliance page — the firm's compliance toolkit (ν. 4557/2018)       */
/* ---------------------------------------------------------------------- */

function CompliancePage({ user, clients, employees, amlRecords, cases, requests, amlFirm, saveAmlFirm, refetchCases, refetchRequests, onOpenCard }) {
  const [tab, setTab] = useState("alerts");
  const f = { ...DEFAULT_AML_FIRM, ...(amlFirm || {}) };
  const setF = (patch) => saveAmlFirm({ ...f, ...patch });
  const empName = (id) => employees.find((e) => e.id === id)?.name || "—";
  const clientName = (id) => clients.find((c) => c.id === id)?.name || "—";
  const alerts = amlAlerts({ clients, amlRecords, cases, requests, amlFirm: f, employees });
  const [caseModal, setCaseModal] = useState(null); const [cf, setCf] = useState({});
  const [reqModal, setReqModal] = useState(null); const [rf, setRf] = useState({});

  // Register (Βήμα 9) rows — one per client, from each client's sheet.
  const registerRows = clients.filter((c) => c.lifecycle !== "lead").map((c) => ({ c, a: amlRecords.find((x) => x.clientId === c.id) || null })).sort((x, y) => x.c.name.localeCompare(y.c.name));
  const counts = { low: 0, medium: 0, high: 0, none: 0 };
  registerRows.forEach(({ a }) => { if (!a || a.status === "not_started") counts.none += 1; else counts[a.riskLevel] += 1; });

  function exportRegister() {
    exportToExcel(`AML_Συγκεντρωτική_Κατάσταση_Πελατών_v${f.registerVersion || "1.0"}_${toKey(TODAY)}.xlsx`, [
      { name: "Κατάσταση πελατών", rows: registerRows.map(({ c, a }, i) => ({
        "Α/Α": i + 1, "Πελάτης": c.name, "Α.Φ.Μ.": c.afm || "", "Νομική μορφή": c.legalForm || "", "Ημ. αξιολόγησης": a?.assessmentDate || a?.completedOn || "",
        "Άθροισμα (Μοντέλο Α)": a ? a.riskScore : "", "Βαθμός (Μοντέλο Β)": a ? a.riskScoreB : "", "Υπερισχύει": a ? (amlAssess(a.riskFactors).forced ? "ΝΑΙ" : "ΟΧΙ") : "",
        "Κατηγορία": a ? ({ low: "ΧΑΜΗΛΟΣ", medium: "ΜΕΤΡΙΟΣ", high: "ΥΨΗΛΟΣ" })[a.riskLevel] : "ΜΗ ΑΞΙΟΛΟΓΗΜΕΝΟΣ", "Μέτρα": a ? AML_DD_LABELS[a.ddLevel] : "",
        "Επόμενη επανεξέταση": a?.nextReviewOn || "", "Έγκριση": a?.approvedBy ? `${empName(a.approvedBy)} ${a.approvedOn || ""}` : "", "Παρατηρήσεις": a ? [a.riskOverrideReason, a.notes].filter(Boolean).join(" · ") : "Χωρίς φάκελο",
      })) },
      { name: "Σύνολα", rows: [["ΧΑΜΗΛΟΣ", counts.low, "Απλουστευμένη δέουσα επιμέλεια"], ["ΜΕΤΡΙΟΣ", counts.medium, "Συνήθης δέουσα επιμέλεια"], ["ΥΨΗΛΟΣ", counts.high, "Αυξημένη δέουσα επιμέλεια"], ["ΜΗ ΑΞΙΟΛΟΓΗΜΕΝΟΙ", counts.none, "Εκκρεμεί φάκελος"], ["ΣΥΝΟΛΟ", registerRows.length, ""]].map(([k, n, m]) => ({ "Κατηγορία": k, "Πλήθος": n, "Μέτρα": m })) },
      { name: "Εκδόσεις & έγκριση", rows: [{ "Έκδοση": f.registerVersion || "1.0", "Ημερομηνία": f.registerDate || toKey(TODAY), "Μεθοδολογία": f.methodologyModel === "B" ? "Μοντέλο Β (0-100)" : "Μοντέλο Α", "Εγκρίθηκε από": f.registerApprovedBy || "", "Υπεύθυνος συμμόρφωσης": empName(f.complianceOfficerId) }] },
    ]);
  }
  function issueRegisterVersion() {
    const cur = f.registerVersion || "1.0"; const [maj, min] = cur.split(".").map(Number);
    const next = `${maj}.${(min || 0) + 1}`;
    const v = window.prompt("Νέα έκδοση της συγκεντρωτικής κατάστασης", f.registerDate ? next : "1.0"); if (v === null) return;
    const by = window.prompt("Εγκρίθηκε από (ανώτερο διοικητικό στέλεχος)", f.registerApprovedBy || empName(user.id)); if (by === null) return;
    setF({ registerVersion: v, registerDate: toKey(TODAY), registerApprovedBy: by });
  }
  function printSelfAssessment() {
    const rows = AML_AUDIT_QUESTIONS.map((q) => { const a = f.selfAssessment?.[q.n]; const ok = q.invert ? a === "no" : a === "yes"; return `<tr><td>${q.n}</td><td>${escapeHtml(q.q)}</td><td>${a === "yes" ? "ΝΑΙ" : a === "no" ? "ΟΧΙ" : a === "na" ? "Δεν συντρέχει" : "—"}</td><td>${ok || a === "na" ? "" : escapeHtml(q.fix)}</td></tr>`; }).join("");
    printHtml("Αυτοδιάγνωση ν. 4557/2018", `<h1>Αυτοδιάγνωση — τα 23 ερωτήματα του ελέγχου ΑΑΔΕ</h1><p class="muted">${escapeHtml(DEFAULT_FIRM.name)} · ${new Date().toLocaleString("el-GR")} · Υπεύθυνος συμμόρφωσης: ${escapeHtml(empName(f.complianceOfficerId))}</p><table><tr><th>#</th><th>Ερώτημα</th><th>Απάντηση</th><th>Τι πρέπει να γίνει</th></tr>${rows}</table>`);
  }
  function printRoadmap() {
    const rows = AML_STEPS.map((st) => { const d = f.steps?.[st.n] || {}; return `<tr><td>${st.n}</td><td>${escapeHtml(st.stage)}</td><td>${escapeHtml(st.title)}</td><td>${escapeHtml(st.out)}</td><td>${d.done ? "✔ " + (d.date || "") : "—"}</td><td>${escapeHtml(d.note || "")}</td></tr>`; }).join("");
    const docs = AML_FILE_DOCUMENTS.map((d) => `<tr><td>${d.n}</td><td>${escapeHtml(d.title)}</td><td>Βήμα ${d.step}</td><td>${f.fileDocs?.[d.n] ? "✔" : "—"}</td></tr>`).join("");
    printHtml("Πορεία συμμόρφωσης", `<h1>Η πορεία των 15 βημάτων</h1><p class="muted">${new Date().toLocaleString("el-GR")}</p><table><tr><th>#</th><th>Στάδιο</th><th>Βήμα</th><th>Παραδοτέο</th><th>Ολοκληρώθηκε</th><th>Σημείωση</th></tr>${rows}</table><h2>Ο φάκελος συμμόρφωσης — 14 έγγραφα</h2><table><tr><th>#</th><th>Έγγραφο</th><th>Παράγεται</th><th>Υπάρχει</th></tr>${docs}</table>`);
  }
  async function saveCase() {
    if (!cf.description?.trim()) { window.alert("Περιγράψτε το περιστατικό."); return; }
    const row = { client_id: cf.clientId || null, detected_on: cf.detectedOn || toKey(TODAY), description: cf.description.trim(), decision: cf.decision || "pending", decision_on: cf.decision && cf.decision !== "pending" ? (cf.decisionOn || toKey(TODAY)) : null, decided_by: cf.decision && cf.decision !== "pending" ? user.id : null, report_ref: cf.reportRef || null, rationale: cf.rationale || null };
    const { error } = caseModal === "new" ? await supabase.from("aml_cases").insert({ ...row, detected_by: user.id }) : await supabase.from("aml_cases").update(row).eq("id", caseModal.id);
    if (error) { window.alert(`Couldn't save: ${error.message}`); return; }
    await refetchCases(); setCaseModal(null);
  }
  async function saveRequest() {
    if (!rf.subject?.trim()) { window.alert("Γράψτε το αντικείμενο του αιτήματος."); return; }
    const row = { authority: rf.authority || AML_AUTHORITIES[0], received_on: rf.receivedOn || toKey(TODAY), due_on: rf.dueOn || null, subject: rf.subject.trim(), client_id: rf.clientId || null, status: rf.status || "open", answered_on: rf.status === "answered" ? (rf.answeredOn || toKey(TODAY)) : null, note: rf.note || null };
    const { error } = reqModal === "new" ? await supabase.from("aml_authority_requests").insert(row) : await supabase.from("aml_authority_requests").update(row).eq("id", reqModal.id);
    if (error) { window.alert(`Couldn't save: ${error.message}`); return; }
    await refetchRequests(); setReqModal(null);
  }

  const tabs = [
    { key: "alerts", label: `Alerts (${alerts.length})` },
    { key: "register", label: "Κατάσταση πελατών" },
    { key: "cases", label: `Ύποπτες συναλλαγές (${cases.filter((k) => k.decision === "pending").length} open)` },
    { key: "requests", label: `Αιτήματα αρχών (${requests.filter((r) => r.status === "open").length} open)` },
    { key: "firm", label: "Στοιχεία γραφείου" },
    { key: "self", label: "Αυτοδιάγνωση (23)" },
    { key: "steps", label: "Τα 15 βήματα" },
  ];
  const sevTone = (s) => s === "high" ? "danger" : "warn";
  const answered = AML_AUDIT_QUESTIONS.filter((q) => f.selfAssessment?.[q.n]).length;
  const okCount = AML_AUDIT_QUESTIONS.filter((q) => { const a = f.selfAssessment?.[q.n]; return a === "na" || (q.invert ? a === "no" : a === "yes"); }).length;
  const stepsDone = AML_STEPS.filter((st) => f.steps?.[st.n]?.done).length;

  return (
    <div>
      <PageHeader title="AML Compliance" sub="Η εργαλειοθήκη συμμόρφωσης του γραφείου κατά τον ν. 4557/2018 — alerts, συγκεντρωτική κατάσταση, ύποπτες συναλλαγές, αιτήματα αρχών, αυτοδιάγνωση και πορεία των 15 βημάτων." />
      <div style={{ display: "flex", gap: 14, marginBottom: 18, flexWrap: "wrap" }}>
        <StatCard label="Urgent alerts" value={String(alerts.filter((a) => a.severity === "high").length)} accent={alerts.some((a) => a.severity === "high") ? C.danger : C.accent} sub="Need action now" />
        <StatCard label="Upcoming" value={String(alerts.filter((a) => a.severity !== "high").length)} accent={C.warn} sub="Within the next weeks" />
        <StatCard label="Clients classified" value={`${counts.low + counts.medium + counts.high}/${registerRows.length}`} sub={`${counts.low} low · ${counts.medium} medium · ${counts.high} high`} accent={counts.none ? C.warn : C.accent} />
        <StatCard label="Self-assessment" value={`${okCount}/23`} sub={`${answered} answered`} accent={okCount === 23 ? C.accent : C.warn} />
        <StatCard label="Roadmap" value={`${stepsDone}/15`} sub="Steps completed" accent={stepsDone === 15 ? C.accent : undefined} />
      </div>
      <div style={{ display: "flex", gap: 4, borderBottom: `1px solid ${C.border}`, marginBottom: 16, flexWrap: "wrap" }}>
        {tabs.map((t) => (
          <button key={t.key} onClick={() => setTab(t.key)} style={{ background: "none", border: "none", cursor: "pointer", padding: "8px 12px", fontFamily: sans, fontSize: 13, fontWeight: tab === t.key ? 700 : 500, color: tab === t.key ? C.accentDark : C.inkMuted, borderBottom: tab === t.key ? `2px solid ${C.accent}` : "2px solid transparent", marginBottom: -1 }}>{t.label}</button>
        ))}
      </div>

      {tab === "alerts" && (
        <Panel>
          <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted, marginBottom: 10 }}>Κάθε υποχρέωση αναφοράς, επανεξέτασης ή καταχώρισης που εκκρεμεί — προς την Αρχή, την ΑΑΔΕ, το ΚΜΠΔ και εσωτερικά. Ενημερώνεται αυτόματα από τους φακέλους των πελατών και τα στοιχεία του γραφείου.</div>
          {alerts.length === 0 ? <EmptyState icon={Check} title="Καμία εκκρεμότητα" sub="Όλες οι υποχρεώσεις είναι ενήμερες." /> : (
            <TableShell headers={["Προτεραιότητα", "Υποχρέωση", "Λεπτομέρεια", "Προθεσμία", ""]}>
              {alerts.map((a, i) => (
                <tr key={i}>
                  <Td><Badge tone={sevTone(a.severity)}>{a.severity === "high" ? "Urgent" : "Upcoming"}</Badge></Td>
                  <Td style={{ fontWeight: 600 }}>{a.title}</Td>
                  <Td style={{ color: C.inkMuted, fontSize: 12.5 }}>{a.detail}</Td>
                  <Td mono style={{ color: a.due && daysUntil(a.due) < 0 ? C.danger : C.inkMuted }}>{a.due || "—"}</Td>
                  <Td right>{a.clientId ? <Btn size="sm" variant="ghost" onClick={() => onOpenCard(a.clientId, "dd")}>Open file</Btn> : a.kind === "firm" || a.kind === "register" ? <Btn size="sm" variant="ghost" onClick={() => setTab(a.kind === "register" ? "register" : "firm")}>Open</Btn> : a.kind === "case" ? <Btn size="sm" variant="ghost" onClick={() => setTab("cases")}>Open</Btn> : a.kind === "request" ? <Btn size="sm" variant="ghost" onClick={() => setTab("requests")}>Open</Btn> : null}</Td>
                </tr>
              ))}
            </TableShell>
          )}
        </Panel>
      )}

      {tab === "register" && (
        <Panel>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: 10 }}>
            <div>
              <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 14, color: C.ink }}>Συγκεντρωτική κατάσταση κατάταξης πελατών — άρ. 35 παρ. 2, άρ. 13 παρ. 9</div>
              <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted }}>Έκδοση {f.registerVersion || "—"}{f.registerDate ? ` της ${f.registerDate}` : " (δεν έχει εκδοθεί)"}{f.registerApprovedBy ? ` · εγκρίθηκε από ${f.registerApprovedBy}` : ""} · Μεθοδολογία: {f.methodologyModel === "B" ? "Μοντέλο Β (0-100)" : "Μοντέλο Α (αφαιρετικό/προσθετικό)"}</div>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <Btn size="sm" variant="secondary" onClick={issueRegisterVersion}>Έκδοση νέας έκδοσης</Btn>
              <Btn size="sm" onClick={exportRegister}>Εξαγωγή Excel</Btn>
            </div>
          </div>
          <TableShell headers={["Πελάτης", "Νομική μορφή", "Ημ. αξιολόγησης", "Μοντέλο Α", "Μοντέλο Β", "Υπερισχύει", "Κατηγορία", "Μέτρα", "Επανεξέταση", ""]}>
            {registerRows.map(({ c, a }) => (
              <tr key={c.id}>
                <Td style={{ fontWeight: 600 }}>{c.name}</Td>
                <Td style={{ color: C.inkMuted }}>{c.legalForm || "—"}</Td>
                <Td mono>{a?.assessmentDate || a?.completedOn || "—"}</Td>
                <Td mono>{a ? `${a.riskScore >= 0 ? "+" : ""}${a.riskScore}` : "—"}</Td>
                <Td mono>{a ? a.riskScoreB : "—"}</Td>
                <Td>{a && amlAssess(a.riskFactors).forced ? <Badge tone="danger">ΝΑΙ</Badge> : <span style={{ color: C.inkFaint }}>—</span>}</Td>
                <Td>{a && a.status !== "not_started" ? <Badge tone={AML_LEVEL_TONE[a.riskLevel]}>{AML_LEVEL_LABELS[a.riskLevel]}</Badge> : <Badge tone="danger">Μη αξιολογημένος</Badge>}</Td>
                <Td style={{ fontSize: 12, color: C.inkMuted }}>{a ? AML_DD_LABELS[a.ddLevel] : "—"}</Td>
                <Td mono style={{ color: a?.nextReviewOn && daysUntil(a.nextReviewOn) < 0 ? C.danger : C.inkMuted }}>{a?.nextReviewOn || "—"}</Td>
                <Td right><Btn size="sm" variant="ghost" onClick={() => onOpenCard(c.id, "dd")}>Φύλλο</Btn></Td>
              </tr>
            ))}
            <tr style={{ fontWeight: 700, borderTop: `2px solid ${C.borderStrong}` }}>
              <Td>Σύνολο {registerRows.length}</Td><Td /><Td /><Td /><Td /><Td />
              <Td style={{ fontSize: 12 }}>{counts.low} χαμηλός · {counts.medium} μέτριος · {counts.high} υψηλός · {counts.none} εκκρεμείς</Td><Td /><Td /><Td />
            </tr>
          </TableShell>
        </Panel>
      )}

      {tab === "cases" && (
        <Panel>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, gap: 10, flexWrap: "wrap" }}>
            <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted, maxWidth: 700 }}>Αρχείο περιστατικών (Βήμα 15): κάθε συναλλαγή που εξετάστηκε ως ασυνήθης ή ύποπτη — είτε αναφέρθηκε στην Αρχή (άρ. 22) είτε όχι (άρ. 30 παρ. 1γ). Η αναφορά γίνεται «αμελλητί» από τον υπεύθυνο συμμόρφωσης. <strong>Απαγορεύεται η γνωστοποίηση στον πελάτη (άρ. 27).</strong></div>
            <Btn size="sm" icon={Plus} onClick={() => { setCaseModal("new"); setCf({ detectedOn: toKey(TODAY), decision: "pending", clientId: "" }); }}>Νέο περιστατικό</Btn>
          </div>
          {cases.length === 0 ? <EmptyState icon={AlertTriangle} title="Κανένα περιστατικό" /> : (
            <TableShell headers={["Εντοπίστηκε", "Πελάτης", "Περιγραφή", "Απόφαση", "Ημ. απόφασης", "Αρ. αναφοράς", ""]}>
              {[...cases].sort((a, b) => b.detectedOn.localeCompare(a.detectedOn)).map((k) => (
                <tr key={k.id}>
                  <Td mono>{k.detectedOn}</Td>
                  <Td>{k.clientId ? clientName(k.clientId) : "—"}</Td>
                  <Td style={{ maxWidth: 360 }}>{k.description}</Td>
                  <Td><Badge tone={AML_CASE_TONE[k.decision]}>{AML_CASE_DECISIONS[k.decision]}</Badge></Td>
                  <Td mono style={{ color: C.inkMuted }}>{k.decisionOn || "—"}</Td>
                  <Td mono style={{ color: C.inkMuted }}>{k.reportRef || "—"}</Td>
                  <Td right><Btn size="sm" variant="ghost" icon={Pencil} onClick={() => { setCaseModal(k); setCf({ ...k }); }}>Edit</Btn></Td>
                </tr>
              ))}
            </TableShell>
          )}
        </Panel>
      )}

      {tab === "requests" && (
        <Panel>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, gap: 10, flexWrap: "wrap" }}>
            <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted, maxWidth: 700 }}>Αιτήματα πληροφοριών από την ΑΑΔΕ, την Αρχή ή άλλη δημόσια αρχή — απάντηση «αμελλητί» (άρ. 22 παρ. 1β). Καταγράψτε κάθε αίτημα με την προθεσμία του ώστε να εμφανίζεται στα alerts.</div>
            <Btn size="sm" icon={Plus} onClick={() => { setReqModal("new"); setRf({ authority: AML_AUTHORITIES[0], receivedOn: toKey(TODAY), status: "open", clientId: "" }); }}>Νέο αίτημα</Btn>
          </div>
          {requests.length === 0 ? <EmptyState icon={Mail} title="Κανένα αίτημα" /> : (
            <TableShell headers={["Παραλαβή", "Αρχή", "Αντικείμενο", "Πελάτης", "Προθεσμία", "Κατάσταση", ""]}>
              {[...requests].sort((a, b) => b.receivedOn.localeCompare(a.receivedOn)).map((r) => (
                <tr key={r.id}>
                  <Td mono>{r.receivedOn}</Td>
                  <Td>{r.authority}</Td>
                  <Td style={{ maxWidth: 320 }}>{r.subject}</Td>
                  <Td>{r.clientId ? clientName(r.clientId) : "—"}</Td>
                  <Td mono style={{ color: r.status === "open" && r.dueOn && daysUntil(r.dueOn) < 0 ? C.danger : C.inkMuted }}>{r.dueOn || "—"}</Td>
                  <Td>{r.status === "open" ? <Badge tone="warn">Ανοιχτό</Badge> : <Badge tone="accent">Απαντήθηκε {r.answeredOn}</Badge>}</Td>
                  <Td right><Btn size="sm" variant="ghost" icon={Pencil} onClick={() => { setReqModal(r); setRf({ ...r }); }}>Edit</Btn></Td>
                </tr>
              ))}
            </TableShell>
          )}
        </Panel>
      )}

      {tab === "firm" && (
        <Panel>
          <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted, marginBottom: 12 }}>Οι υποχρεώσεις του ίδιου του γραφείου (Βήματα 1, 2, 10-12). Οι ημερομηνίες τροφοδοτούν τα alerts: ΚΜΠΔ (άρ. 20), Πίνακας Ι΄ Ε3 ανά έτος, ετήσια επανεξέταση εγχειριδίου, ετήσια εκπαίδευση (άρ. 37), ετήσιος εσωτερικός έλεγχος (άρ. 35 παρ. 3).</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
            <Field label="Υπεύθυνος συμμόρφωσης (άρ. 35 παρ. 3α)">
              <select style={inputStyle} value={f.complianceOfficerId} onChange={(e) => setF({ complianceOfficerId: e.target.value })}>
                <option value="">— δεν έχει οριστεί —</option>{employees.filter((e) => e.role === "admin").map((e) => <option key={e.id} value={e.id}>{e.name}</option>)}
              </select>
            </Field>
            <Field label="Ημερομηνία πράξης ορισμού"><input type="date" style={inputStyle} value={f.appointmentDate} onChange={(e) => setF({ appointmentDate: e.target.value })} /></Field>
            <Field label="Μεθοδολογία βαθμολόγησης (γράφεται μία φορά στην πολιτική)">
              <select style={inputStyle} value={f.methodologyModel} onChange={(e) => setF({ methodologyModel: e.target.value })}><option value="A">Μοντέλο Α — αφαιρετικό/προσθετικό</option><option value="B">Μοντέλο Β — κλίμακα 0-100</option></select>
            </Field>
            <Field label="Καταχώριση γραφείου στο ΚΜΠΔ (άρ. 20)"><input type="date" style={inputStyle} value={f.kmpdRegisteredOn} onChange={(e) => setF({ kmpdRegisteredOn: e.target.value })} /></Field>
            <Field label="Τελευταία μεταβολή δικαιούχων γραφείου"><input type="date" style={inputStyle} value={f.kmpdLastChangeOn} onChange={(e) => setF({ kmpdLastChangeOn: e.target.value })} /></Field>
            <Field label="…καταχωρίστηκε στο ΚΜΠΔ στις (60 ημέρες)"><input type="date" style={inputStyle} value={f.kmpdLastFiledOn} onChange={(e) => setF({ kmpdLastFiledOn: e.target.value })} /></Field>
            <Field label="Έκδοση εγχειριδίου πολιτικών"><input style={inputStyle} value={f.manualVersion} onChange={(e) => setF({ manualVersion: e.target.value })} placeholder="1.0" /></Field>
            <Field label="Εγκρίθηκε στις"><input type="date" style={inputStyle} value={f.manualApprovedOn} onChange={(e) => setF({ manualApprovedOn: e.target.value, manualNextReviewOn: f.manualNextReviewOn || (e.target.value ? toKey(addMonths(fromKey(e.target.value), 12)) : "") })} /></Field>
            <Field label="Επόμενη επανεξέταση εγχειριδίου"><input type="date" style={inputStyle} value={f.manualNextReviewOn} onChange={(e) => setF({ manualNextReviewOn: e.target.value })} /></Field>
            <Field label="Τελευταία εκπαίδευση προσωπικού (άρ. 37)"><input type="date" style={inputStyle} value={f.lastTrainingOn} onChange={(e) => setF({ lastTrainingOn: e.target.value })} /></Field>
            <Field label="Τελευταίο πόρισμα εσωτερικού ελέγχου"><input type="date" style={inputStyle} value={f.lastInternalAuditOn} onChange={(e) => setF({ lastInternalAuditOn: e.target.value })} /></Field>
            <div>
              <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted, marginBottom: 6 }}>Πίνακας Ι΄ Ε3 — συμπληρωμένος για φορολογικό έτος</div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {[TODAY.getFullYear() - 3, TODAY.getFullYear() - 2, TODAY.getFullYear() - 1, TODAY.getFullYear()].map((y) => {
                  const on = (f.e3PanelIYears || []).includes(y);
                  return <label key={y} style={{ display: "flex", alignItems: "center", gap: 4, fontFamily: mono, fontSize: 12.5, border: `1px solid ${C.border}`, borderRadius: 6, padding: "4px 8px", background: on ? C.accentSoft : C.surface, cursor: "pointer" }}><input type="checkbox" checked={on} onChange={(e) => setF({ e3PanelIYears: e.target.checked ? [...(f.e3PanelIYears || []), y] : (f.e3PanelIYears || []).filter((x) => x !== y) })} />{y}</label>;
                })}
              </div>
            </div>
          </div>
          <div style={{ fontFamily: sans, fontSize: 11.5, color: C.inkFaint, marginTop: 10 }}>Οι αλλαγές αποθηκεύονται αυτόματα.</div>
        </Panel>
      )}

      {tab === "self" && (
        <Panel>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, gap: 10, flexWrap: "wrap" }}>
            <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted, maxWidth: 720 }}>Τα 23 ερωτήματα που στέλνει η ΑΑΔΕ στον ελεγχόμενο. Για κάθε ΟΧΙ βλέπετε τι πρέπει να γίνει και σε ποιο βήμα. {okCount}/23 εντάξει.</div>
            <Btn size="sm" variant="secondary" icon={Receipt} onClick={printSelfAssessment}>Εκτύπωση / PDF</Btn>
          </div>
          {[...new Set(AML_AUDIT_QUESTIONS.map((q) => q.group))].map((g) => (
            <div key={g} style={{ marginBottom: 14 }}>
              <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 13, color: C.ink, margin: "8px 0" }}>{g}</div>
              {AML_AUDIT_QUESTIONS.filter((q) => q.group === g).map((q) => {
                const a = f.selfAssessment?.[q.n];
                const ok = a === "na" || (q.invert ? a === "no" : a === "yes");
                const bad = a && !ok;
                return (
                  <div key={q.n} style={{ display: "flex", gap: 10, alignItems: "flex-start", padding: "7px 8px", borderRadius: 6, background: bad ? "#FDECEC" : ok ? C.accentSoft : "transparent", marginBottom: 4 }}>
                    <span style={{ fontFamily: mono, fontSize: 12, color: C.inkMuted, minWidth: 22 }}>{q.n}</span>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontFamily: sans, fontSize: 12.5, color: C.ink }}>{q.q}</div>
                      {bad && <div style={{ fontFamily: sans, fontSize: 12, color: C.danger, marginTop: 2 }}>→ {q.fix}</div>}
                    </div>
                    <div style={{ display: "flex", gap: 4 }}>
                      {[["yes", "ΝΑΙ"], ["no", "ΟΧΙ"], ...(q.na ? [["na", "Δεν συντρέχει"]] : [])].map(([v, l]) => (
                        <button key={v} onClick={() => setF({ selfAssessment: { ...(f.selfAssessment || {}), [q.n]: v } })} style={{ padding: "4px 9px", fontSize: 11.5, fontFamily: sans, fontWeight: 600, border: `1px solid ${C.border}`, borderRadius: 5, cursor: "pointer", background: a === v ? C.accent : C.surface, color: a === v ? "#fff" : C.inkMuted }}>{l}</button>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          ))}
        </Panel>
      )}

      {tab === "steps" && (
        <Panel>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, gap: 10, flexWrap: "wrap" }}>
            <div style={{ fontFamily: sans, fontSize: 12.5, color: C.inkMuted, maxWidth: 720 }}>Η σειρά εναρμόνισης σε πέντε στάδια. Τσεκάρετε κάθε βήμα και, στο τέλος, τα 14 έγγραφα του φακέλου συμμόρφωσης. Ετήσιος κύκλος: κάθε επανεξέταση πελάτη τροφοδοτεί νέα βαθμολόγηση (Βήμα 8) και νέα έκδοση της κατάστασης (Βήμα 9).</div>
            <Btn size="sm" variant="secondary" icon={Receipt} onClick={printRoadmap}>Εκτύπωση / PDF</Btn>
          </div>
          {[...new Set(AML_STEPS.map((s) => s.stage))].map((stage, si) => (
            <div key={stage} style={{ marginBottom: 12 }}>
              <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 13, color: C.ink, margin: "8px 0" }}>Στάδιο {si + 1} — {stage}</div>
              {AML_STEPS.filter((s) => s.stage === stage).map((st) => {
                const d = f.steps?.[st.n] || {};
                return (
                  <div key={st.n} style={{ display: "flex", gap: 10, alignItems: "center", padding: "6px 8px", borderRadius: 6, background: d.done ? C.accentSoft : "transparent", marginBottom: 3, flexWrap: "wrap" }}>
                    <input type="checkbox" checked={!!d.done} onChange={(e) => setF({ steps: { ...(f.steps || {}), [st.n]: { ...d, done: e.target.checked, date: e.target.checked ? (d.date || toKey(TODAY)) : "" } } })} />
                    <span style={{ fontFamily: mono, fontSize: 12, color: C.inkMuted, minWidth: 22 }}>{st.n}</span>
                    <div style={{ flex: 1, minWidth: 260 }}>
                      <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 600, color: C.ink }}>{st.title} {st.q.length ? <span style={{ fontWeight: 400, color: C.inkFaint }}>· ερ. {st.q.join(", ")}</span> : null}</div>
                      <div style={{ fontFamily: sans, fontSize: 12, color: C.inkMuted }}>{st.out}</div>
                    </div>
                    <input type="date" style={{ ...inputStyle, width: 140, padding: "4px 6px", fontSize: 12 }} value={d.date || ""} onChange={(e) => setF({ steps: { ...(f.steps || {}), [st.n]: { ...d, date: e.target.value } } })} disabled={!d.done} />
                    <input style={{ ...inputStyle, width: 220, padding: "4px 6px", fontSize: 12 }} value={d.note || ""} onChange={(e) => setF({ steps: { ...(f.steps || {}), [st.n]: { ...d, note: e.target.value } } })} placeholder="Σημείωση / πού βρίσκεται" />
                  </div>
                );
              })}
            </div>
          ))}
          <div style={{ fontFamily: sans, fontWeight: 700, fontSize: 13, color: C.ink, margin: "16px 0 8px", paddingTop: 12, borderTop: `1px solid ${C.border}` }}>Ο φάκελος συμμόρφωσης — 14 έγγραφα ({Object.values(f.fileDocs || {}).filter(Boolean).length}/14)</div>
          {AML_FILE_DOCUMENTS.map((d) => (
            <label key={d.n} style={{ display: "flex", gap: 10, alignItems: "center", padding: "5px 8px", borderRadius: 6, background: f.fileDocs?.[d.n] ? C.accentSoft : "transparent", marginBottom: 3, cursor: "pointer" }}>
              <input type="checkbox" checked={!!f.fileDocs?.[d.n]} onChange={(e) => setF({ fileDocs: { ...(f.fileDocs || {}), [d.n]: e.target.checked } })} />
              <span style={{ fontFamily: mono, fontSize: 12, color: C.inkMuted, minWidth: 22 }}>{d.n}</span>
              <span style={{ flex: 1, fontFamily: sans, fontSize: 12.5, color: C.ink }}>{d.title}</span>
              <span style={{ fontFamily: sans, fontSize: 11.5, color: C.inkFaint }}>Βήμα {d.step}</span>
            </label>
          ))}
        </Panel>
      )}

      {caseModal && (
        <Modal title={caseModal === "new" ? "Νέο περιστατικό" : "Περιστατικό"} onClose={() => setCaseModal(null)} width={520}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.danger, background: "#FDECEC", borderRadius: 7, padding: "8px 10px" }}>Απαγόρευση γνωστοποίησης (άρ. 27): ο πελάτης και τρίτοι δεν ενημερώνονται ότι εξετάζεται ή υποβάλλεται αναφορά.</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <Field label="Πελάτης"><select style={inputStyle} value={cf.clientId || ""} onChange={(e) => setCf({ ...cf, clientId: e.target.value })}><option value="">— χωρίς / άγνωστος —</option>{clients.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}</select></Field>
              <Field label="Εντοπίστηκε στις"><input type="date" style={inputStyle} value={cf.detectedOn || ""} onChange={(e) => setCf({ ...cf, detectedOn: e.target.value })} /></Field>
            </div>
            <Field label="Περιγραφή — ιστορικό και σκοπός της συναλλαγής (άρ. 16 παρ. 3)"><textarea style={{ ...textareaStyle, minHeight: 90 }} value={cf.description || ""} onChange={(e) => setCf({ ...cf, description: e.target.value })} /></Field>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <Field label="Απόφαση υπευθύνου συμμόρφωσης"><select style={inputStyle} value={cf.decision || "pending"} onChange={(e) => setCf({ ...cf, decision: e.target.value })}>{Object.entries(AML_CASE_DECISIONS).map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select></Field>
              <Field label="Ημερομηνία απόφασης"><input type="date" style={inputStyle} value={cf.decisionOn || ""} onChange={(e) => setCf({ ...cf, decisionOn: e.target.value })} disabled={!cf.decision || cf.decision === "pending"} /></Field>
              <Field label="Αρ. πρωτοκόλλου αναφοράς (αν αναφέρθηκε)"><input style={inputStyle} value={cf.reportRef || ""} onChange={(e) => setCf({ ...cf, reportRef: e.target.value })} disabled={cf.decision !== "reported"} /></Field>
            </div>
            <Field label="Αιτιολογία της απόφασης — τηρείται και όταν δεν γίνεται αναφορά (άρ. 30 παρ. 1γ)"><textarea style={textareaStyle} value={cf.rationale || ""} onChange={(e) => setCf({ ...cf, rationale: e.target.value })} /></Field>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <div>{caseModal !== "new" && <Btn variant="ghost" size="sm" icon={Trash2} onClick={async () => { if (!window.confirm("Διαγραφή περιστατικού;")) return; await supabase.from("aml_cases").delete().eq("id", caseModal.id); await refetchCases(); setCaseModal(null); }}>Διαγραφή</Btn>}</div>
              <div style={{ display: "flex", gap: 8 }}><Btn variant="secondary" onClick={() => setCaseModal(null)}>Άκυρο</Btn><Btn onClick={saveCase}>Αποθήκευση</Btn></div>
            </div>
          </div>
        </Modal>
      )}
      {reqModal && (
        <Modal title={reqModal === "new" ? "Νέο αίτημα αρχής" : "Αίτημα αρχής"} onClose={() => setReqModal(null)} width={500}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <Field label="Αρχή"><select style={inputStyle} value={rf.authority || AML_AUTHORITIES[0]} onChange={(e) => setRf({ ...rf, authority: e.target.value })}>{AML_AUTHORITIES.map((a) => <option key={a}>{a}</option>)}</select></Field>
              <Field label="Πελάτης (αν αφορά)"><select style={inputStyle} value={rf.clientId || ""} onChange={(e) => setRf({ ...rf, clientId: e.target.value })}><option value="">—</option>{clients.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}</select></Field>
              <Field label="Παραλήφθηκε"><input type="date" style={inputStyle} value={rf.receivedOn || ""} onChange={(e) => setRf({ ...rf, receivedOn: e.target.value })} /></Field>
              <Field label="Προθεσμία απάντησης"><input type="date" style={inputStyle} value={rf.dueOn || ""} onChange={(e) => setRf({ ...rf, dueOn: e.target.value })} /></Field>
            </div>
            <Field label="Αντικείμενο"><input style={inputStyle} value={rf.subject || ""} onChange={(e) => setRf({ ...rf, subject: e.target.value })} placeholder="π.χ. Ερωτηματολόγιο 23 σημείων, αίτημα αρχείου πελατών από … έως …" /></Field>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <Field label="Κατάσταση"><select style={inputStyle} value={rf.status || "open"} onChange={(e) => setRf({ ...rf, status: e.target.value })}><option value="open">Ανοιχτό</option><option value="answered">Απαντήθηκε</option></select></Field>
              <Field label="Απαντήθηκε στις"><input type="date" style={inputStyle} value={rf.answeredOn || ""} onChange={(e) => setRf({ ...rf, answeredOn: e.target.value })} disabled={rf.status !== "answered"} /></Field>
            </div>
            <Field label="Σημείωση / τι στάλθηκε"><textarea style={textareaStyle} value={rf.note || ""} onChange={(e) => setRf({ ...rf, note: e.target.value })} /></Field>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <div>{reqModal !== "new" && <Btn variant="ghost" size="sm" icon={Trash2} onClick={async () => { if (!window.confirm("Διαγραφή αιτήματος;")) return; await supabase.from("aml_authority_requests").delete().eq("id", reqModal.id); await refetchRequests(); setReqModal(null); }}>Διαγραφή</Btn>}</div>
              <div style={{ display: "flex", gap: 8 }}><Btn variant="secondary" onClick={() => setReqModal(null)}>Άκυρο</Btn><Btn onClick={saveRequest}>Αποθήκευση</Btn></div>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Root                                                                     */
/* ---------------------------------------------------------------------- */

export default function App() {
  // Browser tab: title and favicon (the ACCO logo embedded above), so the
  // logo shows in the tab without depending on index.html.
  useEffect(() => {
    document.title = "GLORIA";
    let link = document.querySelector("link[rel~='icon']");
    if (!link) { link = document.createElement("link"); link.rel = "icon"; document.head.appendChild(link); }
    link.type = "image/png";
    link.href = LOGO_DATA_URI;
    let apple = document.querySelector("link[rel='apple-touch-icon']");
    if (!apple) { apple = document.createElement("link"); apple.rel = "apple-touch-icon"; document.head.appendChild(apple); }
    apple.href = LOGO_DATA_URI;
  }, []);
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
  const [operatingExpenses, refetchOpex, opexLoading] = useTable("operating_expenses", mapOpex);
  const refetchRental = useCallback(async () => { await Promise.all([refetchRentalPropertiesRaw(), refetchRentalExpensesRaw()]); }, [refetchRentalPropertiesRaw, refetchRentalExpensesRaw]);
  const [billingStatuses, refetchBilling, billingLoading] = useTable("billing_status", mapBillingStatus);
  const [contacts, refetchContacts, contactsLoading] = useTable("client_contacts", mapContact, "name");
  const [notes, refetchNotes, notesLoading] = useTable("client_notes", mapNote);
  const [tasks, refetchTasks, tasksLoading] = useTable("tasks", mapTask);
  const [templates, refetchTemplates, templatesLoading] = useTable("onboarding_templates", mapTemplate, "sort_order");
  const [amlRecords, refetchAml, amlLoading] = useTable("client_aml", mapAml);
  const [ubos, refetchUbos, ubosLoading] = useTable("client_beneficial_owners", mapUbo, "name");
  const [documents, refetchDocuments, documentsLoading] = useTable("client_documents", mapDocument);
  const [firm, saveFirm] = useAppSetting("firm", DEFAULT_FIRM);
  const [amlFirm, saveAmlFirm] = useAppSetting("aml_firm", DEFAULT_AML_FIRM);
  const [amlCases, refetchAmlCases, amlCasesLoading] = useTable("aml_cases", mapAmlCase);
  const [authorityRequests, refetchAuthorityRequests, requestsLoading] = useTable("aml_authority_requests", mapAuthorityRequest);
  const refetchDueDiligence = useCallback(async () => { await Promise.all([refetchAml(), refetchUbos()]); }, [refetchAml, refetchUbos]);
  const [cardClientId, setCardClientId] = useState(null);
  const [cardTab, setCardTab] = useState("details");
  const openClientCard = (id, tab = "details") => { setCardClientId(id); setCardTab(tab); setView("clientcard"); };

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

  // Automatic go-live: when every onboarding task of an onboarding client is
  // done, the client becomes "live" (dated today). Runs for managers only,
  // since only they may update clients.
  useEffect(() => {
    if (!user || user.role !== "admin") return;
    const ready = clients.filter((c) => c.lifecycle === "onboarding" && goLiveReadiness(c, tasks, amlRecords.find((a) => a.clientId === c.id), documents).ready);
    if (ready.length === 0) return;
    (async () => {
      for (const c of ready) await supabase.from("clients").update({ lifecycle: "live", go_live_date: toKey(TODAY) }).eq("id", c.id);
      refetchClients();
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tasks, clients, amlRecords, documents, user?.id]);

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
  if (costHistoryLoading || clientsLoading || feeHistoryLoading || extraFeesLoading || entriesLoading || lockedLoading || leaveLoading || rentalPropsLoading || rentalExpLoading || opexLoading || billingLoading || contactsLoading || notesLoading || tasksLoading || templatesLoading || amlLoading || ubosLoading || documentsLoading || amlCasesLoading || requestsLoading || view === null) return <LoadingScreen />;

  return (
    <div style={{ display: "flex", minHeight: "100vh", background: C.bg, fontFamily: sans }}>
      <FontImport />
      <Sidebar user={user} view={view} setView={setView} />
      <div style={{ flex: 1, padding: "28px 32px", minWidth: 0 }}>
        {view === "timesheet" && (
          <EmployeeTimesheet user={user} clients={clients} entries={entries} refetchEntries={refetchEntries} lockedWeeks={lockedWeeks} refetchLockedWeeks={refetchLockedWeeks} leaveRequests={leaveRequests} />
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
          <AdminDashboard employees={employees} clients={clients} entries={entries} rentalProperties={rentalProperties} rentalExpenses={rentalExpenses} operatingExpenses={operatingExpenses} tasks={tasks} onOpenCard={openClientCard} templates={templates}
            complianceAlerts={amlAlerts({ clients, amlRecords, cases: amlCases, requests: authorityRequests, amlFirm, employees })} onOpenCompliance={() => setView("compliance")}
            billingStatuses={billingStatuses} onOpenBilling={() => setView("billing")} amlRecords={amlRecords} />
        )}
        {user.role === "admin" && view === "timesheets" && (
          <AdminTimesheets employees={employees} clients={clients} entries={entries} refetchEntries={refetchEntries} lockedWeeks={lockedWeeks} refetchLockedWeeks={refetchLockedWeeks} leaveRequests={leaveRequests} />
        )}
        {view === "clients" && (
          <AdminClients user={user} clients={clients} refetchClients={refetchClients} entries={entries} onOpenCard={openClientCard} tasks={tasks} templates={templates} refetchTasks={refetchTasks} amlRecords={amlRecords} employees={employees} />
        )}
        {view === "clientcard" && (
          <ClientCard user={user} client={clients.find((c) => c.id === cardClientId)} employees={employees} contacts={contacts} notes={notes} tasks={tasks} entries={entries} templates={templates}
            aml={amlRecords.find((a) => a.clientId === cardClientId) || null} ubos={ubos.filter((u) => u.clientId === cardClientId)} refetchDueDiligence={refetchDueDiligence}
            documents={documents.filter((d) => d.clientId === cardClientId)} refetchDocuments={refetchDocuments} firm={firm} saveFirm={saveFirm} initialTab={cardTab}
            refetchClients={refetchClients} refetchContacts={refetchContacts} refetchNotes={refetchNotes} refetchTasks={refetchTasks} onBack={() => setView("clients")} />
        )}
        {user.role === "admin" && view === "tasks" && (
          <TasksPage user={user} mode="all" tasks={tasks} clients={clients} employees={employees} notes={notes} refetchTasks={refetchTasks} refetchNotes={refetchNotes} onOpenCard={openClientCard} templates={templates} refetchTemplates={refetchTemplates} />
        )}
        {view === "mytasks" && (
          <TasksPage user={user} mode="mine" tasks={tasks} clients={clients} employees={employees} notes={notes} refetchTasks={refetchTasks} refetchNotes={refetchNotes} />
        )}
        {user.role === "admin" && view === "employees" && (
          <AdminEmployees employees={employees} refetchEmployees={refetchEmployees} entries={entries} />
        )}
        {user.role === "admin" && view === "rental" && (
          <AdminRental properties={rentalProperties} expenses={rentalExpenses} refetchRental={refetchRental} />
        )}
        {user.role === "admin" && view === "opex" && (
          <AdminOpex expenses={operatingExpenses} refetchOpex={refetchOpex} rentalExpenses={rentalExpenses} />
        )}
        {user.role === "admin" && view === "reports" && (
          <AdminReportBuilder employees={employees} clients={clients} entries={entries} rentalExpenses={rentalExpenses} />
        )}
        {user.role === "admin" && view === "compliance" && (
          <CompliancePage user={user} clients={clients} employees={employees} amlRecords={amlRecords} cases={amlCases} requests={authorityRequests}
            amlFirm={amlFirm} saveAmlFirm={saveAmlFirm} refetchCases={refetchAmlCases} refetchRequests={refetchAuthorityRequests} onOpenCard={openClientCard} />
        )}
        {view === "billing" && (
          <AdminBilling clients={clients} billingStatuses={billingStatuses} refetchBilling={refetchBilling} />
        )}
      </div>
    </div>
  );
}
