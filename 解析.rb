require "google_drive"
require 'nokogiri'
require 'open-uri'
require 'date'
require 'CSV'

session = GoogleDrive::Session.from_config("config.json")


#まとめの結果シート
rid = "1faf3-cRzrHLAGVZny8StazPuNT9Tk-yu67jtqbagn5Y"
rws = session.spreadsheet_by_key(rid).worksheet_by_title("シート1")

count = 0

#対象のシート
ids = [
  "1-ugNKh0YccuDeNN4feDoX0B1h5Q9t-YZuJdnhBdzTik",
  "1EunsTvEjsNQmS6xuwfvRFBgPIVkOGqu2Hdru-nFhx0A",
  "1Y_1QiOjgHiVHFwir0HlYVMJusYLJYDW_t7vgy8_E0qs",
  "1cDKeQ4VtULaFQ_PO62RHWMccU5XGa6sfmyR2oJouq4Y",
  "1QrfWCYDgATH8CqU3oPH6wYA6eJ8wlGBwaxItTM6NvK8",
  "10Yl3QNBncl722WTocZWIeX-H97wvpK_nYuby0bDsl08",
  "1940523pPY0Kebdx-R7IqZw31eBtYzsMD1c3b_Wo5lfg",
  "1vuyxutMDEFFhPssetl8onYsiQV7v94n3gLLUa4EKZYc",
  "11WfByUVUu8sUPTCuz8iNZK7HeuCQOQw14jFZ2DkXs-Y",
  "1f6dlbm8gzJ6Tzg77bZhRZL5FcK4xI2sq5jyxJHQf6SI",
  "1ic8Tepdrr9tslpMbOa-4OE6FAP-QE_9782FfpC2hSXw",
  "16FKjtRtYjuNu_eY8r9eORFL-bxvomEIKPYWrNEGPo-I",
  "1uSzMnTCc5PyCCgCyccfB_HGESm6GH-mDSLYJihGI764",
  "1PdpwARUijDd7RpUO-Sd21RutHagAhDAXESDePa4BtYA",
  "1qXc_lduIQCoOfA8QIiD2qhYDhF8_PnbzK13y07FH6iA",
  "1i-QKxC6lzz6PB4p3W6JjiDsC1Xu2mzasl_qyuWZUwXI",
  "1bcgfHEIgAThyinJ1bBa6lTcUGW9fxw5YtWx0POdfH_4",
  "1tu_i3gIs0d87O96yzirANZHAQA-rEl7mXX5Kf84CjSk",
  "1q3mgskBoB4HS6W8qHalhJCiqmkpJzMgEgqweK21Amm4",
  "1ldKrYNh1HcbmgLCGJDbGUTd_L0tDOwrQd3Riti2FzB0",
  "11O0C0aOD_l0MyMMSh4yjcJng_peS-9Fye2Dp11yFv8Y",
  "1mASO7jhsn7A035ruGkupIsA1aNZpRKcjzn3PnupXAWU",
  "1WU6p994b4TKq6IdGCWiDg847qEWnqlqMVwH9G4lJs6E",
  "1mKZbxO2wewS3WKLiRN6ynOfl7XM92o6B0GcG-mHlj0s",
  "18q9gFG0urAdfdGgzbMH9g9MC4Ij7G4iqMYoM1v9cK30",
  "1GL7zEOvBjRgsCVtT2tcsnqSJCEDowEnaVAqsKLQru24",
  "1e2a7aVqdR9KHcLjKSpv5kZLYu7KR_Huq-VLYuBNZ2U4",
  "1NRIExqudWjfPK5fx2_mcr8OLXpDadnuWLD_PUDMo9X0",
  "1zgVtpyLLLPztlOzL5grIfD-QL84nJnUz99Rn92FI2Fk",
  "1fqw4dtg2lJSxGAj-Gh6N_12evB6D-WWHl928pRyq61U",
  "1oJotugotGrq0VajN-4hN2vTIFbKYx0DrDvHS8-xgb04",
  "1jeYb3XO6ZUz5MSwJX8S4tdZKq7CenUhT-GpoqO462_E",
  "1nEN0FTY1JMm01l8IA1tnwthlNom-_3S23n4ga1zqBkU",
  "1tyaKgWXSPugd_D20AbBnYL_T-Yv7eu7ilcxcwhCSG1M",
  "1ETTrRCFLl1VzWsds5M4E9Sp1LEfJNUFsingdGasw6Qg",
  "1nHZq0f5FDBDbqoZgwplODn6qLtsgyLcKC_FQt5uGqV8",
  "1TQLat7fyn1ywwKHUyR4VQn_eMkpNRwvtPwhwPvUD7qw",
  "1bkU_7iLph4re43BFd3Gj5LePPhdm4-fySqhJb1QWtHM",
  "1kKjFUvmkdQ4L0Xkjt5jX9RD7H9Cj49lcFtDskoxMfqM",
  "1uLsUPJbvB8j6IMBlcBMTbKECGPeMqKSDWArsSpPqov4",
  "1erFeyu_DNyxwdychs29-WDXpU96W4BhIK7QxMjc4rCI",
  "1kmvNEVsIc3ppBslwtqM7CdFfOmiFWA3UzdzqvIApzbY",
  "1xHO7Rs3wcBSJMLbod-SpzqbCs7C6SoAwhELNDMa03Hs",
  "1guT17p6fOeCKD_OrJwA1vRzuI_kgtd7Z2_PflnlCgD8",
  "1nb7vAUsN17-8ax_eswYXQPhQ2pTgDaO5SPhN-6FKye4",
  "1lVppOfDaiAUZh7CGYLzsHFJFM_k1PZJrDlQ7RZKaYwA",
  "11iDRJLd3gwHxAkn4MZuSkMIxc-SCFIcg4xJG7ja7Ims",
  "1oHUcWCUQYB1yCYLifKAhh2wxrIQ6xTkVR4FJSHAFvek",
  "1j8CgoAUXPnZ0rI-kTIHFmiwSFSJ0uMRsVnp7-oymiJk",
  "1UAUCE5ybxMTzwyAiranI5-FEP3iNcz0uN6mFY4MQA9Q",
  "18wj4K5DNbeqBE_DyENRHTsv8f8A8r_ztsSuzDGWS76o",
  "1TJjM3JxxSDdKF_YiSfi_CkeF-pX0bvduHpi_4eDV3U4",
  "1rZZ2nH3np0PAZ-9eIL11r_75nJp6AG3zG6bAKz-Rpmw",
  "1iGTuEB7EgJ3jU7ALacrVjlvR0o3pwYH9_pi54-76x9k",
  "1HiNIHPKPYx60m251PxBBCwIRJj7Xd3VIK2ncV7eC3cw",
  "1uCz-qXn-nWtUb5WkmBHainLsnZOSeqWMEDwHpOTAKxQ",
  "1R6_uX_slOducTMtrH1-GnhmHpbnN-tAEjGN35czv6uw",
  "1LDQkf3-9oXk6dDk0zxylxaFr64JdE0J6uBZUB9dhvT4",
  "1rnYIM4Se-VSb2jl12zfLCIPLd4-KMuiLxCkzA4NxPxg",
  "1giJfcrHZ3FpJCoeiTQVKq2960OZu-FiFMHM8FX8RyFw",
  "1oURcy065DLE4R4USmyhtlFTcNFjoDLpp4VjwDtXrivA",
  "19SVs8ranWEv5AckrLwIj9W-3v69Jqg2Gu-2nSp60A6w",
  "1lM8QlduAkj5cVUtseRQAkOfZsqC8PLFkZxvon9e9bpw",
  "13Z9pugChAGTQDbHlR6g7CTeevRbs4NtmhosmLh9h588",
  "1l8x3DXqK3ffYGUPjuOqDTAN3SHRwaT90tCDVsBE1grA",
  "1dZVvRi0dmXJtsPOJ7ST3vDFvjTXPzgMvVERPBRl_7tY",
  "1vn_Df6OsV4VYAnXDxNivwyIDqzhJoft6Zh5KjO2d-DE",
  "1wdRqPKfhzdkf04MxPCJktMn2dbqrd5gsMr2bKKh13X0",
  "1MJHjOt73fJ2KYceo4jEJbsDve7iGGMD1zahKVBb6iz4",
  "1mtAcahKtPeP4gGCsXuq4PipTtMl7-036Qhw8JaKo3ns",
  "1EKshIkUEERem1welpJKlcWTlNVTdyvgt10ZmNQCH5Gk",
  "1SK8xJUPABHc5QsqB9E1wzgljY32cSRvDwxVMQqyUgh4",
  "1ALcU7gy95ZmlZfC7jX_u9B1rsI4Ty4Bh5PEAdeJzf9M",
  "1ZNbdFSun8WyXSMjRQRI8odmh31q85HSoCNUMk0Ea2Wo",
  "1chYPiKfQJAQ62I2sSpbypoLE3cSsyZNEk-QVkO6t5uE",
  "16bmugbhmqdnbByZNMjTac0mTsMDI_sfXAulT6SdHISQ",
  "1ySSzgZc4fP4LfNurFLL69mdgd_ZRYc2nMcKv7yeIyyc",
  "1yioeofzJBCXxzBAfCOVOQw0H_PJI3p_9Jy3wsu4NExE",
  "1EPcyxHb5j43Xb5LrL7MabsVIHhkMI5Pw6p_n2RmS9q0",
  "1411aeGmN4OoIgjt0Y6Jf4Ucm5PL7nketAxPgjjVg2Wc",
  "1uaHAOkOWDiG1O4_Xu-KMxm2yX0FEbxug6_zHwNAhP34",
  "19TfkoSIpGEdIkPHrUvF1jp4XVIyJGCTh-Qvs-Xd4qaI",
  "1h8dijRlMwZ6iv4iwcwAACELgx-HpYvMy8THdWxhIi6g",
  "1s6p56sgulqptnxfWR4eBcxBd57IGaBtWF-faolYu-8w",
  "1ADj-K1D7a0PnEzdX7dwoBV3npAemaln5VDpAbf99qBI",
  "1r-En3Jb3fSEfa4PcP6bCMvdJfEcbIm2yvVOi47r_S30",
  "1tbFhBWfBMO0L5L3OoZgCbqmgfv-esQh3F2UhWGz-a04",
  "1Q80EOWvmB1z7g-JT5j2-_2rRetM30vagEGmTYpZQjw8",
  "1dL6NC5f8zh1o8GaK-QuNvo9Q6odAqnvySyrQdvgH3Dc",
  "1juLtGS3daO2jxPsktzwSUhgcc9g5XXTC6v2UXAjm-Ok",
  "17MWdFWfIn_C3VKamJkjdlDjhCMqYA71IQfqiqIkIjBE",
  "1-OOgJ5zEBcQLyn3VbQ2g281m4L_eSfZHLfizKB9csCU",
  "1oc6fSp9EAxEgaEDfdCmj6zuv1P-eclyevQQySxgs3hQ",
  "1Nz-gpy_ayOVW_kzrQ0O6ZrWCDr21Dkk9SytshdSpSqU",
  "1W3uhKj22T0mflZVGJRCkeyHydBHNRZrjAXthUiSGsaw",
  "1d9LPucosH-qrw3a5KqpOi04O45WnQhVrQLPCo_0RS2Y",
  "13NuQQQIQSxjhFQcVf_rNY6nahYfYkE8-Rqbv2vxqbZ0",
  "1swAxADyiVSGXGtRdZUvocKzbGeoXcWrwBqdDkk-TdWM",
  "15GKTIsO7a3ZkgzT6FS0SbZXQQbbNnYG-3e7bGXdevvE",
  "1xXb1betYqXnjJQxaobTL_-WLM-01L_c7-FPLZicewIM",
  "1IsCyddIqpSqdgXTi3Cf-Ksy5UUqFtk7HKGSa1A19IH0",
  "1YOVk1eKFaJgb1m5gyIUM9PV43HaENVSlLxWOP_Kzojw",
  "17srvYPu5CZtJD4KIZo0Q6I3hgaRv5yTtYi7JcTx62yo",
  "1mGoFSd_Pau2IhnP3XwNM_1JF4FBX1DgPCkyUEgNG-PM",
  "1ge1pkpFd1gA5bK9JAX2ILRZXH1h62ty7O6l3ECUsf-E",
  "1hMyX3Z294q7Vo1h13mlpUf190XHl37NhVxji9laoUXg",
  "14fcjfnf7GEJD0bc91JqsQklCfK2JwCXYFH_P0bqKkbA",
  "1l705cWqebw_9igZxqWXvRLWB9_PCUuuRf7hyBF1XH5U",
  "1Omrfh9mAi-t2fbMV12RZdcKftzgRyMnoo4_EtWfIsIs",
  "1lNO2wPBzsLrUAP3pmQx7XxVJZXdH4MlYL3B6J6_ZCaY",
  "1f2Zt7EUSiJlMsUzfh3YBCqjeWRrCJ6Q5FleeTvulcvw",
  "1p23QMLEXz07wMXkZqr8xcOEW8ZDXFxHAA72SrhWvWMA",
  "1ch9u6yjsI4j4Swa6Dcmu685k4hFLBG7mwhRi44y46Hw",
  "1hYKnhLrYVIH8GgcK3DjBBzeqTTXcxTrcAf311BwEY_w",
  "1pnAHrZkMnUGh7ffHLyUUXa-fdsvTYbT3J5cRcmrr8NE",
  "1jaGNoeeDPrjUM2wB3u6pZ8PYzsrOwhxQwj0QojNn2kc",
  "16-m2dg1SGmAkZcqLkyOnoiPu0XjjaIxT_kkEgB5iYq8",
  "1Fn5-9tnC-bfSGCe7zcqfszlzvc6d277RXMt8p4OGMYk",
  "1HJBq07BAMRS_CoSOFhpPN-uuPIbS-Ku90sSoKo36D14",
  "1vZ1fBTLMZMlRsh7f4JjW83DUoxLV3ZMdaRTCWd8FEzw",
  "1Ns4LQlky8SwKqLOnUVYsAUo_S1Mr9sWPDQQ1oi4FMI4",
  "1dn_XZa9gT47MaiB6IwUgfZghtDspvZuyujnu7UlWNxA",
  "1KckNCzaN8dYYePlj1mpEV2nXiT_hef059sMsTjfWlp8",
  "1TZ0Boga9f_o7WCJiPk4FP4AmWeG9xFBrUDT6f25L4hA",
  "18jh2rVGC0-b7e1jL29FCMxDsh2vxnhh5x1E3MEuSxkg",
  "1DZzcxstX_DHFmarPEeFC4f7-8XVjy_6JxXoHOvG6sk8",
  "1eIgDJEy4nDqr-wBPuzom3L0f12XFSf56y7VVdGy2_5I",
  "1mcl6QgfsIjFFWHeLdU9awBNveX-ctZxlY6cbrqbXbdE",
  "1dxPhM3LaBmX75aX463bOrntOEwYZFC8PJ1X4vGN75MM",
  "1V-BEUciEdNb8GYBJhRWYJIO57IggNus1Osyn8YsOewU",
  "19CMcgnlLH_QGdEzymtpFV1OLuedfANp4cADgr9Ba1iw",
  "1rowpn7wLnI0leC7c51T6mUDeJlfcn1VHf3vgbGjCYiU",
  "1kZb3TUhHC0f67WrFtWzShGHnMdRfgkz1TC1DWr-pqEA",
  "1lsta7maqX1jX-8TQOC8yQyOr7HoP_L_aUKiMuKzdcmE",
  "1OQg7fWGem6tGTGPdE0I1JCJCZMPW6A-Pekzc7wDMqNU",
  "1L6LeJ6cMXdi8802LnejNUDR-GjmV29-wgDYVtd9Ge4o",
  "1uth_YnnsoPnyT8Fuj11Tc7AmGKk96PRYJ38AMXhzMIg",
  "18TCIdtLsygP6yJk1FErIVPGDsGxWdVxPzJxDgdsdgTQ",
  "1rJV9J9Sfjdp9kHktmBFrUWf9vYbYctHF4C7SNA66KCI",
  "1Po9YYeRAsLeGgLAzeKSu_V_HGe1ZgPDVn1tZH-_KO0w",
  "1Mq3FT3E3xtD72i-fEaUWYMbAOvLQVHEFlgX1fToRpaI",
  "1LEw3PBf9VEPlMvK3KPMvnCTwmx4G7SqTU6355dWk7Mo",
  "1BnTYsHVqu46BYH0FjtgpRlI3bcq2AP730BDgJr0k9tE",
  "1NZP_zPAFWhUDGpDnexpvCN4UEu8DqCLAFxEuQVzwHEg",
  "10ERctofgkqJ7DyKQ3dXgfsjkqJZBuEjnoNUEI7KaX6c",
  "1bk2l_3GYb7JClomzVUu_whOJdgvo-jHe6c-3xnFLCs8",
  "1JRrjgapzxmUyGJW1Sjaiem-HQHScDYUOSOx53N0fG5U",
  "1_C66ogjd8yNVTSW0IDIOyxUrmTV5rHYKSSYyseHttHc",
  "1lNqHqXPbQYsXknw1QzWW_VFnZdEeV-Nfx9d9V34x7G8",
  "14RQ_Jr5J0xhU8c8MrctIYUv7VBCE9Ec9TXbxmXW0YMY",
  "144Q231IiMSCN_qFtP8MoJ0fXbiHR7dITQjf65UW4SHo",
  "1xFIuMOFjd0q7i0ScEoENNvcVoMHSNtVefkOy06NsPmc",
  "1-CFCV2dCsS13TcAiwXT87j6vJZr12HS9H3T58YBqG_s",
  "1wVTlPdAT2dJPD9MxULDQraJ2D09A41sU2UeQ663hwUs",
  "1Mcz3w1q4g97uCYXgOxMlJwF1u_TCr1VR_0shufVWmsM",
  "1ueJ-QX6F2rYBHm4N_dzHQvovqvY7Sil2I9LSCqlmxdE",
  "1fM8mExYEaGsl0lsDSeQThQBM3WmiF7CbzNr49lt0QdM",
  "1fsYFoOjcGI4o-BAj_3K6QVkpFArU5IDvPm6W9lDIioA",
  "1LUo7TWvHRYcJS3eReWOUH5QZAGKNJEDsQ34UYsxrySw",
  "1aws6rVYNQik1nmp4doa6C-21cWY6YqKufGOrhbwgxq0",
  "1NT98lOf8bBHodo24f5nY_EM_yuNB51t7xhVfsm_McxU",
  "1vDmcyeiqEyuoZP407FJYmfRqDtwzEoTAxs072mJ5Zwk",
  "1FfzpFvBDa0SbdRxW1qeCWZgMdI4r2-togSyujoHVMUU",
  "1wRxK0ckQvSJbp4nEkSygcxqt4MV00rX-4CStrWbjRtQ",
  "1LHo_gc8-MOWaXg_WV1y9Z4_xE9EDn8gwFeRQSHiyLIs",
  "1HD7HeCNLzPX1LEBeo-Dq4vdfqbbUwT8kzQheEAf7Tn4",
  "13QaCeoS1btbqQtQULIOXdTOJRnQy2B_YHloclfCDFpA",
  "109JJssOl1C_oAp3R_XR7AK4g0mam0oyYqcFKtCFRoj4",
  "18mlaRvuTYj2x9vN1d0BGXVyWw8tArhPz1wtJ5EU64RY",
  "1vqlYNPESs2y3CC1EfHQw8Bb9uubBYWM_QXYfuafvbv8",
  "1k5ysv0XC56QIhgB9cVbhgJzJ1osagPnqVoxcCVxyd8c",
  "1E0u5Xpmd4lFm6tDLUsyPsfmBYyBk1JyNINcN7pbY2uc",
  "1o48KDp5oN2CpJM3bZMuAlcannyl4jEthatVGdqP9Zp8",
  "1eZZeYYnWQJ9WUnsCFBhdojPeyC8k4Ih_sY4mpbUJFEw",
  "1Cw2BnuG4thXRSpcZcD9FTFBGd4Xff-o0OfQHyLwD2fw",
  "1wQp4bMCnK5oqKMYGnVcVS2HY9LtkPAe5Ta6bNAYxYuQ",
  "1rN_t_vYeIDbeN9UOmO-b7zQwLM_8XZLlrHIDnLO3HJ4",
  "1ZreQ2U3vtg9cwjLgzakAC0oriGqet3sPJ4qyuYltrzU",
  "1dCQAWI61LHn07rGIRohcunfxpq3MbO3S0vwb2tlqaD8",
  "1a0-YTn7ZyK3Kj0s5nauYf3JHaV7Gfxkj287QsRHn4QQ",
  "1qNRfr54XtI71UhzvT1B9kmaRK6Y8O24YEiEXS71ehHw",
  "1HLR3NV9BJUemZrup7tzhDpctE7A72kpZfCdFqoK4BlM",
  "1R2-9Yi2Sjx43uIxvo0Gc3HRihvPkCyYSbxN38j93CtM",
  "1H_3qAs4CVu4unm_NC9PSkbuKja4XpTsBGYg_Sb2oq00",
  "1dJMdCZWhzxduN-Jm8aaUZYXzDHIvLqEBBH7KJ31z0Yg",
  "1K-RvQtxWr8WGi_YoPtXmKXe_TR6FlQw1r2w3rHbhhwo",
  "1eM-b7l8OtCmoPWjEF2BO1uX4ZWWflNbU_nnSsTW1xV4",
  "1AD_MFSsPBmDGpYH04QCFa8Nm8GnuO8_qTMW28VQA6Mg",
  "10OivLiVBM9CooRN6nFt11Zs8HUbqMHmNWV6vw70htVw",
  "1KZYj5f4srGMf4zQQfDriqmfCcAgV2dfUqjsG7bJLCHY",
  "17Deeoe7dDdk6j8bpRkiCyylMsL2HivTFBtKA4MNa3L8",
  "1r0E31B39LyDJJ1uwam1-RPgpT29DZnYP0JltJzdxp5o",
  "1WkoLwONYmbFP8e7A-r4ZsDr4CP5LsdULL0mc5OsvpDc",
  "1An492sWP29GZWHHrb2fBPSYF-hE7HQJZIWoL6PI2xcI",
  "1JuNqyovIaafBQFKfS4jP4_fmY_vt1tCimBHznIogIpw",
  "1BqOOyGa_KwjLLL5qg4KIJu6kdrJpeldNEVDGtbuBYKk",
  "1ck2KCFYgs-HlmBvAfe671-i5qpf6RptMF9CcED9pGFU",
  "1Lm8imZ07be8-IjZluopprYTlCrVwLO7rQsb6yG7Zqhc",
  "10bEKBd3y6X9jV_BH_Z3gSljq43p050ctvhKh1qiDC9g",
  "1uQ-JnoqL8FAlBWfLVI3AKHXVAO7JwKDlf4eclXXbJEo",
  "1PZRNn0bHxCnrTD7nRhS5NXUR904e-ggcNiQl2obaUMk",
  "1Is397xY7ig7pJXGTwkh_2G9UMH3haB4uaICA9VGNT9A",
  "14wInWquCqa2tstQvnOMEECYVvrZr6wNmFjJxQ_KRhfw",
  "1_1IqxTB1PnW8cT39zXZqG5VJvFTL1px_hlVffMBkmio",
  "1408KitEiV6MUM2TS6fAW0naWJecHV16mECj3O1TrluU",
  "15d20mFIZa3XTTka5Sh_G2Xu74lyMBraJMnaWjhvGzvw",
  "1-cGL4vTLUZsfyjbumbHMtUUKg0RTqUk_FjMPnlFwacs",
  "1t3uq0oTfsiYcykIwVzWlqmZYC7i_5CbXaEh3DyXYXMY",
  "1uhMF13jei_E51g5bR-R05s0MHfZZg3vAu49WhAJ8uos",
  "15ebDZ2JoExgERqmKPMCIt6qZ_FKHp0aREjVkvTd0RFM",
  "1vfYjf6m9PvIl_OTkS67xXZGmrDadMxJO5HGpRLESKXY",
  "1PoNm1trgjtGym3pu1Q6tuiAoSZxdV4sE2cCZ8ONr_po",
  "10hZuOiFJfty_MiYuez31RB5SgP-PnLtCppRUZcAZIeo",
  "1qEVwBFPMoIo9spbG-CBR7BK-WkfsCDpWxlp9tyI4AjI",
  "119Y-6WexadAQz3oBgSyuC59dRGb3v6ECujmab_BSkeo",
  "1MxBhzva9ISSSra2k9F8dta4AunQ24B-92vHbgZAZqtU",
  "1-gGivXfJbjleD7sNz2eskJ74oAaaz_Be7oUcPB3aiq4",
  "1pCxW2kdMdZ2c4M9XKEYduLsf8xskigUSP4ALxLY3yoA",
  "18xQo22aTHxr9vzHnGwIlpiXWvPIOJMY4scnbkuCwl0w",
  "1_r5gGLltsktPmnHw_t_i7RWJpZAWckoN7TU1n1RZdTY",
  "1QTTplAHgRatYnyMmviAX1pzacbNtQSYx792jGHCNFXo",
  "1zXlREd0gQIzalHfmIuQhVk6zbin640a02ocip6ZIxLA",
  "1N0xPJ2tvSM-ySjOATkLi27ZTZ7zNvgaBo7Of6pL7l8s",
  "1msfbJt5vwrmbLBt0G6As-iy-ISvN0m8Bw-xO0d1wNQE",
  "1DAH9GJoIZkCOk_D4Nsy8RLgvPVH_TX4RzJwMuJZisBo",
  "18Gi9JzZ_vG2Sq01dVzT-VYR1g08S7JquQ7-fC4y5Oz4",
  "1OwF-OvZJYYu0Z8Z39yN571umECF3r2uAwQwcb4Sq6Tc",
  "1xR5fo7iDTHqbx7h3KlsxzfGX_muHIR1ApLtVZF7A_mw",
  "1hDk5pB92rk8HxVqi7YvOcIO2517IHB81b1yn72Do7hg",
  "1Z7bxER20RR2CAmI5z1xISfrKUu3uvldlizMqsZdOVXA",
  "1IYSVOreMqBO1ZT8lCi_5fXvcWsJOMSwBTvJazOSqFk4",
  "1V2Z_fI-9ShODkCZ7ROrD6llsf7Fq5NLO5fVXNYhoc_4",
  "1zZtDiVMnFWJ-bhQLEEXd3rveObXKwlPXSKAI6eeSLt4",
  "1G8Wr2h-u4bR6XFezCGcyLjNCs3Vg9ZhYN4KPTakELzE",
  "1RVOHjRMk1wYBvdDjY-6Tn4515eASQZIXOOPI5je4QqQ",
  "1i2qe4RE-NZ9FUqtPQK6Rsq1qwiRfJg3C1L-9McYEU7w",
  "1SKJFq54wJy-cNFSxjpzH6hp_h5HMKcywb77PaTrGR0k",
  "1XrlIuSb6uIn8iCXW73-nNNOuV74LuDP9vNPuJndHj5U",
  "1pbCPwaZ8fkuqrWFR9H3JT0E3mW-vqBFoaC-5fPSsDio",
  "1rn78leBzkFQQNZ0MwJ0h3GW7deJT4W7o9KHMrmC-boM",
  "1-eVNVtjaoY5DHA0Gdb1gNe_c5IyLvswewb43B1FqeRM",
  "1X1zykSUK8m2eUryB1vOJGFNR5N4Xthhns-iffh9LMH8",
  "1rWrRUHSIA-BxYCvTPIbt1mejpJjQDulit9UYKFnv-Mo",
  "12mYg7bWq0_UEt4m5_wqcM9cm3zvgUduh0bNcsPWId8U",
  "1jqQL94LMqM5YiwerTwsgzUldDPDPnHkJDz8VEkGycXQ",
  "1VCD186sqsSN9MrvhQbcAabAEkcENzuHLlampu3Ic5rQ",
  "1LeW8utWU8dUgsm3xs2Y4THkB-eGxuTvuBX9UjbJJHPQ",
  "1hp6MV1R6ucBX1jr-K2Ex_6nAM8WhWxS0A6rEv0Q74RY",
  "16M3BncvCFLJx7Htv3CVhVTEc82c2VNDIuIdyI5iIQSM",
  "1RD4qYWuFg7WDWI3n2B-2ebWfY9npm1_BWResNKc2jnM",
  "1fc4zIOhb8uybobQBrVnWh7zFNqBQkEfF1hgkmHllmvI",
  "1Xr9tyPWDWlSYlbXDYf3VPlMAaDYmYtzm4htwBRb7icU",
  "1L4AifALUZqMyjLJgFwOYnbm23oVwjNbCLlhl8xvwJ30",
  "1CqAX-Grx_ztn79JPj6TaHtGaOqGquFA0o8bqVsR08wM",
  "1oUzXGr4jK2hFymfxOpYRvKalRRJtu3bVPzs4srKoCPs",
  "1wmZEdgduVIrTlq5UzcT2-rL4XFnuWLEQRajgn_tm9AY",
  "1Ia879rJcZ_CRmNL0qLB052qDVQe0N1b_tlW-_cRwkic",
  "1uQcMT3D_fEWxXPsVy_X-i5bgRa5mOahTIB7b9b7_qvQ",
  "1_VIJlFebrDmWSwd_YYdNJrZMsvEGSMEMWYlt8HjBTd8",
  "12LG8NR-ROBiMIFm7ydyaciFJJlYVo_ClkABs8pa6S_g",
  "1UKuNWAaT9q1b-HfWjsEHLvyNDGfAPNskx3oFMqBc8PM",
  "12do7G6iDepAz01VgbKvrqz9ZH1OdaNmuArAWUO7-_s4",
  "1nF2CU5uR3Yv3ycxBd5mCpkNSMGhPeS-RIzQpzmAmmB4",
  "1bxNntC97iTSrA13_6ahnDBOburp1tzLA-zb3-g6gn4U",
  "142hIvECv64qR-Rbc0aOOclJcuBTGU0QV7hztIY2GWhU",
  "15KXEeWJicyeJ-7n5sqob7iYCVSFD7BeaxaDAE4YrwZY",
  "1E3kUND7hv6G3kwiLyJ5cTMR6ajwg5hg4CYuVMG3dN6s",
  "1GyNFF2M5IpKQY6qHR8F0GBa0gct6Z9LNWaUG_xZd17g",
  "1ZuRJ4Gxf6XXfi_VC-aEn3_p_4SbWYCKeHb7tvnb0piM",
  "1ojI3lS5t4fugvC8UPLgqWml95PX1S9Fg8pIqvkqR9_k",
  "1YDX55KzeZGGkJxEbdvyafwpsl_4aYelHiy8FR8JoXeE",
  "1eDQ0iIH5yM1dm11KWwC_23n0JeWWQRI-bR52HFo9JLI",
  "1gvezracrxQNzKlu1-pFMDGGGPWzM-0lGcTYu7e5miJU",
  "1uJTsynGgfMqCTS8YPwKJoA_4mRPSzKJGuOo1M-xkKfU",
  "1BmPOSZI8Jl1nzT8FvduvooH-X-onX7jGO049bszh8Ig",
  "1JTxGSvKb2FAvug0lQprGlSZ0cTGV3fsPQsK5gQeY5n8",
  "1JSJw333rl7oJgP9kKdzBTCJOcJj1Rcg6oh9rWONf4Z4",
  "1QMvCT5lZxzcNAm58G-wQDGJnM_6mhDNbGnm6N1hPwDw",
  "1MAMfuaNbNGEakwTjf9ua7lGtBddpQg78zuI2RXPVBqw",
  "1uLRjdsQ8oCkFoAnmkSbPTmfqM0a61EBuKsnhNHDvKiY",
  "1NIpn4SzCpkFXXrs3EkAC27aqC-zc-TVg6IhG5qUDGPs",
  "146OMuLuyCTul8UyXZisBkDZ8NxE2OV20NKmqyc_U0FI",
  "1117c2PxZyy-SPYFcwyxVcGk_ULzu9ZlLrjnf35OUor0",
  "1Xb5StagluE3sDRpHGB9L2J0s_H-h5eUzNrECmNHi0FE",
  "1LL2wNxYPOqwAwGBriWifkC1fRp8Hk9evHO5EGB0qTB0",
  "1sKeeI-_BiJPvpEtlrm9moFqmWoFhk0sYImhIcluZB_U",
  "19VmD3ReIAFzO5JkjxMYzBgqilzI4jpqrUyiWD7V-GpM",
  "1EGlR_dcVL1xXA3OWvgs4QKaPCAdiSfGVCRIjn1vCx7E",
  "1ElOs58UuHHD_uaXuMt4wIYa7f-HP9fZ1ksWNvoYSmJ0",
  "1LVNhLwFyaNbiEEZxzHF22HAKQDZRB8iQHi12PW4HrrQ",
  "1SVDRTFv3dGvxMPjFwtC9HXbL7FGg8nNi3P6F5qBmi9w",
  "12pzhfifk00ctg9zomhwkJYP3A0Wxm14SrbtnxShOBqQ",
  "1aTej88m2uprRuhbm03gyVbz_bjGvdGJzDX73GQF0Tgc",
  "1ngZTHIa00JVCuCUXT3csbnG3MVA1_zXgS9JfjeH2Kco",
  "1AHN1_LweZPYgR-gk50vUfijUPJedvFdVBCONwvVmdL4",
  "1BOaY3Xcb4tJSfjams4jxVvBrYFN9kZKFncCxu0eYz3c",
  "11jwYBa0azaKoVKTYVz9e_mnN1lRdZROWti9OkqUc-e8",
  "1mF8IVldJNrS6OsLjhCVxi-nfOe5VCPVp-o0YPbedR3M",
  "1OpKNQy0NoU7UUVwOgvT379JgjQ8g3PbVHjwSYNEF71Q",
  "1vY7iP7Lf81QtE2OES9hPdORlW9GzKEtumTDq6K2xOuI",
  "1ASoEPMUClu-xQ8rnIezRNlD9I90zIQckHutp2q2hIB0",
  "1ey2hlL_sk8a8OyyZ7UU8-LxVnhaiJ8jh6VE7iev5YyY",
  "1LQia64mSBs8Oz3301pUvRlXcKnds8sgfR4Mlpo9tPRI",
  "1TFReED7R76gKtLS_TSJXLOiX7SR-SAXmS3X82425wC4",
  "19yZF9JERy5HoujKYR4GfeO1OKT0XW_U-6TZWulIOpfg",
  "1jPByVtVomstFeGwIbVns76WpXAX1HA4NhUpQyigwqdY",
  "15j_KkN5cvsH-5F40IqG-8H2l3KoW8ZySz-PTnToWp9I",
  "1EfSgvOw1ZRAqWLm0tzxo-MvyjvArD1s6Vxphc5in-Ig",
  "1HiP_eMUHTYg-twnPsZQyIq32KeLjznaq99YFV0uyvgQ",
  "13pOelfILB7d5rYHe2XIqhp7p66B_D_vriU1_-HMH_s0",
  "1fUJvEBy_FKTgALlCAVNCwonpW697ZO2bQ6H5fD8YP8c",
  "1KHEVkKxY-zmkUdHOK-mpi5t2SHS-6oq3N0oGAxhZnwA",
  "1M7ChTaUZEW3zaJSoQI3y9976kU_1-2glX8BYHXHlReU",
  "1YLzt4U8rR2py_Mu_hA7oJJOsZ5lMYhZF8XayxhxxgUw",
  "1UMjsGwV1856iY8N9easGFpP5u0Ddq-MyvurK2VH7KMs",
  "1zbf4-dHXshHuovkbBypWj-PUYjkKzHxYLnagwPduQC0",
  "1IyXuwJQB3xeMk_J1-lCC2m0Au3DVREVSyoyVttZ75dM",
  "1AoCijKaK7q0fMmtnKDd-Bu_zpeAqgFcOQmRxmRENqxk",
  "1bztDUJ19rybW15bsVPlIeS7WHxX1ETg88pT1Qv8DtXw",
  "1mpL__r858rErj7gV-HbtW2ZI2ot-oMM2f4UlE2S0xMc",
  "1fT40rGZrZfyez7dpQEtcyPbzTDuwKVAHEoCBA4Vt_Cs",
  "1dq522zj5-x26zDLYQgVX-YpfkWM6Cqt8AmmM4cfxa7o",
  "1bnsKMds789amZEWVh45d8-de1HppHRJsqCus5B5WhAU",
  "1M3B6NFMj0GFyHVJnTzVM4bRp-Uo6K6jP7J4k-8lI1dA",
  "1M7zqC6qlAkuJP5N7fgbl0fI6l3QTE1SCSWhZuxLGY_Q",
  "1ffx8UvciRHE9yZTGLiqSzADlihxjBH9EkmO-Xs21MZI",
  "1kyKmVzY85W3G_OPS7wTUirFc1TSh4_IMgT7Sjn0vnvc",
  "1bnsnzJTEGshXUza2k6BQoqPdWZmPEArJnii3mnp-iYs",
  "1mxNgtGB9BDVmAQ12G6ZK0CieQu6XDNMFsPvHJcHpKI4",
  "12enFA7pml2W4fYtTucuu0hxPR9DpAA2wBG2ZBfP74Lo",
  "1-NTTxLjgAjAWK8GY0ie7VB-rfFjjwFDNU2Z4kuXgDtg",
  "1HsDiv99jTb9HA5yaJN9zuG7rdgcF7fZ0xFVaGs_I0H4",
  "12WTExifgr90vavo0HGi__S7Tve98MwlJEzXJGJpIqzU",
  "1GOiE0RzdgLV_t_mNG58ACPtmfzKHZytNbgQX_biB9wQ",
  "15byhIjpsm_CesXxnW8gDYJrl3Hz6tVzBw-o6UJDwdFQ",
  "1k6S8E3LncoW-CMVUPxKfjagv8dyU-5tmf8Ml_-Apeg0",
  "1rEv8oun7plULwGbSX3PLAOcAeFi1U3G3HQR4YTbZdnc",
  "1fkJnIjwv4TkfIDxVBT8fMs6GqwppSB7dOW7F2RAPo7E",
  "1jbKHh-TSyDRl08q4yeLntlV3xuxzHRvZcGgrVRuiRvg",
  "1Ip1Igjj-nG_pUrWyouSZx-wh2zNI7JiDFdqLJdul35I",

]

ids.each{|id|

  p id

  ws = session.spreadsheet_by_key(id).worksheet_by_title("シート1")

  #今現在の最終行を取得
  for n in 3..220
    b = ws["A#{n}"]
    p b
    if b != "" then
    	rws[n + count,1] = ws["A#{n}"]
    	rws[n + count,2] = ws["B#{n}"]
    	rws[n + count,3] = ws["C#{n}"]
    	rws[n + count,4] = ws["D#{n}"]
    	rws[n + count,5] = ws["E#{n}"]
    	rws[n + count,6] = ws["F#{n}"]
    	rws[n + count,7] = ws["G#{n}"]
    	rws[n + count,8] = ws["H#{n}"]
    	rws[n + count,9] = ws["I#{n}"]
    	rws[n + count,10] = ws["J#{n}"]
    	rws[n + count,11] = ws["K#{n}"]
    	rws[n + count,12] = ws["L#{n}"]
    	rws[n + count,13] = ws["M#{n}"]
    	rws[n + count,14] = ws["N#{n}"]
    	rws[n + count,15] = ws["O#{n}"]
    	rws[n + count,16] = ws["P#{n}"]
    	rws[n + count,17] = ws["Q#{n}"]
    	rws[n + count,18] = ws["R#{n}"]
    	rws[n + count,19] = ws["S#{n}"]
    	rws[n + count,20] = ws["T#{n}"]
    	rws[n + count,21] = ws["U#{n}"]
    	rws[n + count,22] = ws["V#{n}"]
    	rws[n + count,23] = ws["W#{n}"]
    	rws[n + count,24] = ws["X#{n}"]
    	rws[n + count,25] = ws["Y#{n}"]
    	rws[n + count,26] = ws["Z#{n}"]
    	rws[n + count,27] = ws["AA#{n}"]
    	rws[n + count,28] = ws["AB#{n}"]
    	rws[n + count,29] = ws["AC#{n}"]
    	rws[n + count,30] = ws["AD#{n}"]
    	rws[n + count,31] = ws["AE#{n}"]
    	rws[n + count,32] = ws["AF#{n}"]
    	rws[n + count,33] = ws["AG#{n}"]
    	rws[n + count,34] = ws["AH#{n}"]
    	rws[n + count,35] = ws["AI#{n}"]
    	rws[n + count,36] = ws["AJ#{n}"]
    	rws[n + count,37] = ws["AK#{n}"]
    	rws[n + count,38] = ws["AL#{n}"]
    	rws[n + count,39] = ws["AM#{n}"]
    	# rws[n + count,40] = ws["AN#{n}"]
    	# rws[n + count,41] = ws["AO#{n}"]
    	# rws[n + count,42] = ws["AP#{n}"]
    	rws[n + count,43] = ws["AQ#{n}"]
    	rws[n + count,44] = ws["AR#{n}"]
    	rws[n + count,45] = ws["AS#{n}"]
    	# rws[n + count,46] = ws["AT#{n}"]
    	# rws[n + count,47] = ws["AU#{n}"]
    	# rws[n + count,48] = ws["AV#{n}"]
    	rws[n + count,49] = ws["AW#{n}"]
    	rws[n + count,50] = ws["AX#{n}"]
    	# rws[n + count,51] = ws["AY#{n}"]
    	rws[n + count,52] = ws["AZ#{n}"]
    	# rws[n + count,53] = ws["BA#{n}"]
    	rws[n + count,54] = ws["BB#{n}"]
    	# rws[n + count,55] = ws["BC#{n}"]
      rws[n + count,56] = ws["BD#{n}"]
      # rws[n + count,57] = ws["BE#{n}"]
      rws[n + count,58] = ws["BF#{n}"]
      # rws[n + count,59] = ws["BG#{n}"]
      rws[n + count,60] = ws["BH#{n}"]
      # rws[n + count,61] = ws["BI#{n}"]
      rws[n + count,62] = ws["BJ#{n}"]
      rws[n + count,63] = ws["BK#{n}"]
      rws[n + count,64] = ws["BL#{n}"]
      rws[n + count,65] = ws["BM#{n}"]
      rws[n + count,66] = ws["BN#{n}"]
      rws[n + count,67] = ws["BO#{n}"]
      rws[n + count,68] = ws["BP#{n}"]
      rws[n + count,69] = ws["BQ#{n}"]
      rws[n + count,70] = ws["BR#{n}"]
      rws[n + count,71] = ws["BS#{n}"]



    else
      count = n + count - 3
      break
    end

  end
  rws.save
}
