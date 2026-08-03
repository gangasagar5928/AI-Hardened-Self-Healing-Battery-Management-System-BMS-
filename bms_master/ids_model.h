#include <string.h>
void add_vectors(double *v1, double *v2, int size, double *result) {
    for(int i = 0; i < size; ++i)
        result[i] = v1[i] + v2[i];
}
void mul_vector_number(double *v1, double num, int size, double *result) {
    for(int i = 0; i < size; ++i)
        result[i] = v1[i] * num;
}
void score(double * input, double * output) {
    double var0[4];
    double var1[4];
    double var2[4];
    double var3[4];
    double var4[4];
    double var5[4];
    double var6[4];
    double var7[4];
    double var8[4];
    double var9[4];
    double var10[4];
    if (input[1] <= 3.7032474279403687) {
        memcpy(var10, (double[]){0.0, 0.0, 1.0, 0.0}, 4 * sizeof(double));
    } else {
        if (input[0] <= 24.5864839553833) {
            if (input[2] <= 157.90762992203236) {
                memcpy(var10, (double[]){0.0, 1.0, 0.0, 0.0}, 4 * sizeof(double));
            } else {
                memcpy(var10, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
            }
        } else {
            memcpy(var10, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
        }
    }
    double var11[4];
    if (input[2] <= 372.8379821777344) {
        if (input[2] <= 171.68343353271484) {
            if (input[0] <= 236.77948600053787) {
                memcpy(var11, (double[]){0.0, 1.0, 0.0, 0.0}, 4 * sizeof(double));
            } else {
                memcpy(var11, (double[]){0.0, 0.0, 1.0, 0.0}, 4 * sizeof(double));
            }
        } else {
            if (input[2] <= 347.6702117919922) {
                if (input[0] <= 22.412516117095947) {
                    memcpy(var11, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
                } else {
                    memcpy(var11, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
                }
            } else {
                if (input[0] <= 24.00372314453125) {
                    memcpy(var11, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
                } else {
                    memcpy(var11, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
                }
            }
        }
    } else {
        if (input[2] <= 378.5174102783203) {
            if (input[3] <= 1.2305285334587097) {
                memcpy(var11, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
            } else {
                memcpy(var11, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
            }
        } else {
            memcpy(var11, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
        }
    }
    add_vectors(var10, var11, 4, var9);
    double var12[4];
    if (input[2] <= 362.76478576660156) {
        if (input[1] <= 5.001250505447388) {
            memcpy(var12, (double[]){0.0, 0.0, 1.0, 0.0}, 4 * sizeof(double));
        } else {
            if (input[3] <= 0.5284878313541412) {
                memcpy(var12, (double[]){0.0, 1.0, 0.0, 0.0}, 4 * sizeof(double));
            } else {
                if (input[2] <= 351.4104766845703) {
                    if (input[0] <= 20.88770580291748) {
                        memcpy(var12, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
                    } else {
                        memcpy(var12, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
                    }
                } else {
                    if (input[1] <= 13.763655662536621) {
                        memcpy(var12, (double[]){0.75, 0.0, 0.0, 0.25}, 4 * sizeof(double));
                    } else {
                        memcpy(var12, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
                    }
                }
            }
        }
    } else {
        if (input[0] <= 28.572508811950684) {
            memcpy(var12, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
        } else {
            memcpy(var12, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
        }
    }
    add_vectors(var9, var12, 4, var8);
    double var13[4];
    if (input[2] <= 373.1455383300781) {
        if (input[0] <= 4.377608597278595) {
            memcpy(var13, (double[]){0.0, 1.0, 0.0, 0.0}, 4 * sizeof(double));
        } else {
            if (input[1] <= 5.132729530334473) {
                memcpy(var13, (double[]){0.0, 0.0, 1.0, 0.0}, 4 * sizeof(double));
            } else {
                if (input[2] <= 351.98443603515625) {
                    if (input[0] <= 22.412516117095947) {
                        memcpy(var13, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
                    } else {
                        memcpy(var13, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
                    }
                } else {
                    if (input[0] <= 29.006959915161133) {
                        memcpy(var13, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
                    } else {
                        memcpy(var13, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
                    }
                }
            }
        }
    } else {
        if (input[2] <= 378.5174102783203) {
            if (input[2] <= 376.4996643066406) {
                memcpy(var13, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
            } else {
                memcpy(var13, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
            }
        } else {
            memcpy(var13, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
        }
    }
    add_vectors(var8, var13, 4, var7);
    double var14[4];
    if (input[1] <= 3.568259835243225) {
        memcpy(var14, (double[]){0.0, 0.0, 1.0, 0.0}, 4 * sizeof(double));
    } else {
        if (input[2] <= 367.3971252441406) {
            if (input[1] <= 77.68099403381348) {
                if (input[1] <= 9.293649673461914) {
                    if (input[0] <= 27.102972984313965) {
                        memcpy(var14, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
                    } else {
                        memcpy(var14, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
                    }
                } else {
                    if (input[0] <= 22.387789249420166) {
                        memcpy(var14, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
                    } else {
                        memcpy(var14, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
                    }
                }
            } else {
                memcpy(var14, (double[]){0.0, 1.0, 0.0, 0.0}, 4 * sizeof(double));
            }
        } else {
            if (input[1] <= 18.003426551818848) {
                memcpy(var14, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
            } else {
                memcpy(var14, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
            }
        }
    }
    add_vectors(var7, var14, 4, var6);
    double var15[4];
    if (input[2] <= 373.1455383300781) {
        if (input[0] <= 4.822446644306183) {
            memcpy(var15, (double[]){0.0, 1.0, 0.0, 0.0}, 4 * sizeof(double));
        } else {
            if (input[0] <= 267.62273597717285) {
                if (input[0] <= 23.28746271133423) {
                    memcpy(var15, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
                } else {
                    memcpy(var15, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
                }
            } else {
                memcpy(var15, (double[]){0.0, 0.0, 1.0, 0.0}, 4 * sizeof(double));
            }
        }
    } else {
        if (input[1] <= 17.62642765045166) {
            if (input[1] <= 14.377916812896729) {
                memcpy(var15, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
            } else {
                if (input[2] <= 394.8004608154297) {
                    memcpy(var15, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
                } else {
                    memcpy(var15, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
                }
            }
        } else {
            memcpy(var15, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
        }
    }
    add_vectors(var6, var15, 4, var5);
    double var16[4];
    if (input[0] <= 25.878528594970703) {
        if (input[0] <= 2.6785529255867004) {
            memcpy(var16, (double[]){0.0, 1.0, 0.0, 0.0}, 4 * sizeof(double));
        } else {
            memcpy(var16, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
        }
    } else {
        if (input[1] <= 5.008707404136658) {
            memcpy(var16, (double[]){0.0, 0.0, 1.0, 0.0}, 4 * sizeof(double));
        } else {
            memcpy(var16, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
        }
    }
    add_vectors(var5, var16, 4, var4);
    double var17[4];
    if (input[0] <= 24.512977600097656) {
        if (input[1] <= 75.23401927947998) {
            memcpy(var17, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
        } else {
            memcpy(var17, (double[]){0.0, 1.0, 0.0, 0.0}, 4 * sizeof(double));
        }
    } else {
        if (input[1] <= 5.891427397727966) {
            memcpy(var17, (double[]){0.0, 0.0, 1.0, 0.0}, 4 * sizeof(double));
        } else {
            memcpy(var17, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
        }
    }
    add_vectors(var4, var17, 4, var3);
    double var18[4];
    if (input[0] <= 24.5864839553833) {
        if (input[2] <= 157.8972818478942) {
            memcpy(var18, (double[]){0.0, 1.0, 0.0, 0.0}, 4 * sizeof(double));
        } else {
            memcpy(var18, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
        }
    } else {
        if (input[1] <= 6.34586238861084) {
            memcpy(var18, (double[]){0.0, 0.0, 1.0, 0.0}, 4 * sizeof(double));
        } else {
            memcpy(var18, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
        }
    }
    add_vectors(var3, var18, 4, var2);
    double var19[4];
    if (input[1] <= 3.568259835243225) {
        memcpy(var19, (double[]){0.0, 0.0, 1.0, 0.0}, 4 * sizeof(double));
    } else {
        if (input[1] <= 14.24989366531372) {
            if (input[0] <= 24.753026962280273) {
                memcpy(var19, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
            } else {
                memcpy(var19, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
            }
        } else {
            if (input[2] <= 112.22734732925892) {
                memcpy(var19, (double[]){0.0, 1.0, 0.0, 0.0}, 4 * sizeof(double));
            } else {
                if (input[2] <= 398.5419616699219) {
                    if (input[2] <= 359.63197326660156) {
                        memcpy(var19, (double[]){0.0, 0.0, 0.0, 1.0}, 4 * sizeof(double));
                    } else {
                        memcpy(var19, (double[]){0.1, 0.0, 0.0, 0.9}, 4 * sizeof(double));
                    }
                } else {
                    memcpy(var19, (double[]){1.0, 0.0, 0.0, 0.0}, 4 * sizeof(double));
                }
            }
        }
    }
    add_vectors(var2, var19, 4, var1);
    mul_vector_number(var1, 0.1, 4, var0);
    memcpy(output, var0, 4 * sizeof(double));
}
