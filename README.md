# Visit Summary Generator

## Table of contents
* [General info](#general-info)
* [Screenshots](#screenshots)
* [Technologies](#technologies)
* [Setup](#setup)
* [Features](#features)
* [Status](#status)


## General info
Python script used to scrape patient information from the schedule page of Cerner's EMR to print out a personlized summary page for the provider's notes for each patient visit. Built using python 2.7, tkinter, docx, pywinauto and win32api. This was a great project because I was able to learn graphics programming with tkinter, how to interact with different windows, how to inject information into a word document and how to turn a python script into an executable for distribution.

## Screenshots
![image](https://user-images.githubusercontent.com/6406075/117045059-512ad700-accc-11eb-9c05-0c021303c492.png)

![image](https://user-images.githubusercontent.com/6406075/117045572-e1691c00-accc-11eb-856c-6740c4e8ff2e.png)


## Technologies
* python 2.7
* tkinter
* docx
* pywinauto
* win32api
* cx_freeze
	
## Setup
Clone the repo, install dependencies and run setup.py to generate the binary. 

## Features
* Custom dates
* Automatically calculates the next business day
* Enter custom days off for business day calculations
* Easy to use GUI
* Works with Cerner's EMR PowerChart and posibly others
* Easily customize

TODO:
* Allow customization of providers

## Status
Project is: _in progress_

[comment]: <> (_finished_, _no longer continue_ and why?)
