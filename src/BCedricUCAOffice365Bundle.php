<?php

namespace BCedric\UCAOffice365;

use BCedric\UCAOffice365\Service\UCAOffice365;
use BCedric\UCAOffice365\Service\GraphAPI;
use Symfony\Component\Config\Definition\Configurator\DefinitionConfigurator;
use Symfony\Component\DependencyInjection\ContainerBuilder;
use Symfony\Component\DependencyInjection\Loader\Configurator\ContainerConfigurator;
use Symfony\Component\HttpKernel\Bundle\AbstractBundle;



class BCedricUCAOffice365Bundle extends AbstractBundle
{

    public function configure(DefinitionConfigurator $definition): void
    {
        $definition->rootNode()
            ->children()
            ->arrayNode('uca_api')
            ->children()
            ->scalarNode('url')->end()
            ->scalarNode('login')->end()
            ->scalarNode('password')->end()
            ->arrayNode('graph_api')
            ->children()
            ->scalarNode('client')->end()
            ->scalarNode('tenant')->end()
            ->scalarNode('client_secret')->end()
            ->end();
    }

    public function loadExtension(array $config, ContainerConfigurator $container, ContainerBuilder $builder): void
    {
        $container->import('../config/services.yaml');

        $container->services()->set(UCAOffice365::class)
            ->public();

        if (array_key_exists('uca_api', $config)) {

            $builder->autowire(UCAOffice365::class)
                ->setArgument('$url', $config['uca_api']['url'])
                ->setArgument('$login', $config['uca_api']['login'])
                ->setArgument('$password', $config['uca_api']['password'])
            ;
        } else {
            $builder->autowire(UCAOffice365::class)
                ->setArgument('$url', '')
                ->setArgument('$login', '')
                ->setArgument('$password', '')
            ;
        }

        $container->services()->set(GraphAPI::class)
            ->public();
        if (array_key_exists('graph_api', $config)) {
            $builder->autowire(GraphAPI::class)
                ->setArgument('$clientId', $config['graph_api']['client'])
                ->setArgument('$tenantId', $config['graph_api']['tenant'])
                ->setArgument('$clientSecret', $config['graph_api']['client_secret'])
            ;
        } else {
            $builder->autowire(GraphAPI::class)
                ->setArgument('$tenantId', '')
                ->setArgument('$clientId', '')
                ->setArgument('$clientSecret', '')
            ;
        }
    }
}
