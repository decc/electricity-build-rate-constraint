# Quality Assurance of Electricity Build Rate Constraint Model

Author
  : Tom Counsell

Date
  : 9 September 2013

Canonical source
  : http://github.com/decc/electricity-build-rate-constraint

## Aims of the model

The model's purpose is to explain the consequence of delaying construction of low carbon generation.

The consequences that are shown are cost, CO2 emissions and build rates of low and high carbon generation.

## Overview of the model 

The core of the model is a very simple annual model of the UK electricity system from now to 2050. It deals with variability of demand and supply by numeric convolution rather than half hourly dispatch. It is implemented in Excel. It allows assumptions about costs, build rates, demand, demand shape and evolution of costs over time to be varied.

## Model design

## Key assumptions and rag rating

| Assumption        | Source                      | RAG | Notes                                   |
| 2012 cost of gas  | Guesstimate                 | Red | Need to match to DECC fossil fuel price |

## Worry list

* Add known decomissioning profile
* Inputs: Divide into three phases rather than using fixed dates
* Need to validate the shape of my load curve
* Need to think of a way of dealing with failiures and reserve
* Validate assumption for min and max as proportion of mean
* Update capacity figures to 2012
* Need to add availability factors to other generation types
* Add estimate of marginal price (need to think about pricing when market is tight)
* Add storage
* Finance costs?
* Add a more realistic profile for legacy plant (especially nuclear) shutdown.
* Check plant life-times
* Energy security
* Allow today's costs of technologies to be varied
* Allow emissions target to be varied
* Allow CB4 level to be varied
