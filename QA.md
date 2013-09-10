# Quality Assurance of Electricity Build Rate Constraint Model

| Author           | Tom Counsell     |
| Date             | 9 September 2013 |
| Canonical source | http://github.com/decc/electricity-build-rate-constraint |

## Aims of the model

1. The model's purpose is to explain the consequence of delaying construction of low carbon generation.
2. The consequences that are shown are cost, CO2 emissions and build rates of low and high carbon generation.
3. The model is built in a single sheet of excel, in order to be easy to follow

## Overview of the model 

The core of the model is a very simple annual model of the UK electricity system from now to 2050. It deals with variability of demand and supply by numeric convolution rather than half hourly dispatch. It is implemented in Excel. It allows assumptions about costs, build rates, demand, demand shape and evolution of costs over time to be varied.

## Model design

### Overview

The model assumes there are three generation types:

1. Low carbon intermittent generation - currently, this is something like onshore wind
2. Low carbon dispatchable generation - currently, this is something like nuclear. To represent CCS the emissions factors for this would need to be changed.
3. High carbon dispatchable generation - currently this is like gas, although the emissions factor for it is varied a little to represent the proportion of capacity that starts off as coal.

### Demand

The model assume that demand can be modelled as a probability distribution function whose shape is a square distribution next to a triangle distribution. i.e., there is no chance that demand is below a certain level. There is an equal chance that it is above that level but below the median. There is a high chance it is at or a bit above the mediam. Then there is a rapidly declining chance that it is higher, falling to zero at a maximum. This shape is specified by a minimum demand, maximum demand and mean demand.

### Intermittend generation

The model assumes that intermittent generation has a probability distribution with zero as its minimum output, falling linearly to its mean output which has a probability equal to its load factor, then linearly to zero.

The output of intermittent generation is assumed to be independent of the variation in demand.

### Demand net wind

The model carries out a numerical convolution of the demand pdf and the intermittent generation pdf in order to get a pdf of demand net wind. The numerical convolution is done over discrete intervals of 1 GW so has some error.

### Merit order

Intermittent generation is assumed to dispatch first, then dispatchable low carbon generation, then dispatchable high carbon generation.

### Capacity lifetime

In any given year, the model assumes that (1/lifetime) of the existing capacity is shut down. In addition, capacity disposals can be manually specified in the model (so as to represent the scheduled shutdown of the existing nuclear fleet). The maximum of the manually specified or the calculated shutdown is taken.

### Security of supply

Each of the three generation types has a 'guaranteed availability at peak' factor which is the proportion of the the capacity that can be relied upon. There is also a 'capacity margin' factor. The peak demand is multiplied by the capacity margin to get the total capacity required. Existing capacity is derated by its guaranteed availability at peak. The difference is then filled by building new high carbon generation capacity.

## Key assumptions and rag rating

| Assumption        | Source                      | RAG | Notes                                   |
|-
| 2012 cost of gas  | Guesstimate                 | Red | Need to match to DECC fossil fuel price |
| Guaranteed availability at peak | Guesstimate   | Red | Need to get from National Grid or OFGEM capacity reports |
| Capacity margin | Guesstimate | Red | Need to get from National Grid or OFGEM report |
| Legacy capacity lifetiem | Guesstimate | Red | Need to get from National Grid or OFGEM report? |

## Worry list

* Is the NPV wrong because it doesn't value capacity after 2050?
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
